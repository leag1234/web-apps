/*
 * ElixirOffice: Word's Layout Options button and object anchor (document editor).
 *
 * When one picture, shape or chart is selected, a small button appears next to its top
 * right corner, as in Word. It opens the Layout Options flyout:
 *   - In Line with Text;
 *   - With Text Wrapping: Square, Tight, Through, Top and Bottom, Behind Text, In Front of Text;
 *   - Move with text / Fix position on page (floating objects);
 *   - See more... (the Layout dialog, Position tab).
 * When the selected object floats, the anchor of the paragraph it is attached to is drawn in
 * the margin, with a lock when the anchor is locked.
 *
 * The wrapping choices go through the ribbon's Wrap Text menu (same history point, same
 * position kept when an inline picture starts floating). Geometry comes from the sdkjs API
 * asc_GetSelectedDrawingLayout (pixels of the document view).
 *
 * Licensed under the GNU AGPL v3 like the rest of web-apps.
 */
define([
    'core'
], function () {
    'use strict';

    var esc = function (s) { return Common.Utils.String.htmlEncode(s === undefined || s === null ? '' : String(s)); };

    // Word-like layout icons, 32 px tiles drawn on a 20 x 20 grid: text lines follow the text
    // colour, the picture uses the accent of the icons
    var PIC = function (x, y, w, h, extra) {
        return '<g class="elx-lo-pic"' + (extra || '') + '><rect x="' + x + '" y="' + y + '" width="' + w + '" height="' + h + '" rx="1"/>' +
            '<path class="elx-lo-mount" d="M' + (x + 1) + ' ' + (y + h - 1) + 'l' + (w * 0.3) + ' ' + (-h * 0.45) + 'l' + (w * 0.22) + ' ' + (h * 0.25) + 'l' + (w * 0.14) + ' ' + (-h * 0.15) + 'l' + (w * 0.24) + ' ' + (h * 0.35) + 'z"/></g>';
    };
    var LINES = function (rows) {
        return rows.map(function (r) { return '<rect class="elx-lo-line" x="' + r[0] + '" y="' + r[1] + '" width="' + r[2] + '" height="1.2" rx=".6"/>'; }).join('');
    };
    var ICONS = {
        inline: LINES([[2, 3, 16], [12, 13.8, 6], [2, 17, 16]]) + PIC(2, 7, 9, 8),
        square: LINES([[2, 2, 16], [2, 5, 4], [14, 5, 4], [2, 8.5, 4], [14, 8.5, 4], [2, 12, 4], [14, 12, 4], [2, 15.5, 16]]) + PIC(7, 4.5, 6, 9),
        tight: LINES([[2, 2, 16], [2, 5, 5.5], [12.5, 5, 5.5], [2, 8.5, 4.5], [13.5, 8.5, 4.5], [2, 12, 5.5], [12.5, 12, 5.5], [2, 15.5, 16]]) +
            '<g class="elx-lo-pic"><circle cx="10" cy="9" r="4"/></g>',
        through: LINES([[2, 2, 16], [2, 5, 6], [12, 5, 6], [2, 8.5, 18 - 2], [2, 12, 6], [12, 12, 6], [2, 15.5, 16]]) +
            '<g class="elx-lo-pic" opacity=".85"><path d="M10 4.5l4.5 4.5-4.5 4.5-4.5-4.5z"/></g>',
        topAndBottom: LINES([[2, 2, 16], [2, 4.5, 16], [2, 15, 16], [2, 17.5, 16]]) + PIC(6, 6.5, 8, 7),
        behind: PIC(5, 4, 10, 12, ' opacity=".55"') + LINES([[2, 3, 16], [2, 6.5, 16], [2, 10, 16], [2, 13.5, 16], [2, 17, 16]]),
        inFront: LINES([[2, 3, 16], [2, 6.5, 16], [2, 10, 16], [2, 13.5, 16], [2, 17, 16]]) + PIC(5, 4, 10, 12),
        anchor: '<path class="elx-lo-anchor" d="M10 2a2 2 0 0 1 .75 3.86V7.5H13v1.5h-2.25v6.4c1.8-.3 3.3-1.6 3.8-3.4h-1.3l2-2.5 2 2.5H16c-.6 2.9-3.1 5-6 5s-5.4-2.1-6-5H2.75l2-2.5 2 2.5h-1.3c.5 1.8 2 3.1 3.8 3.4V9H7V7.5h2.25V5.86A2 2 0 0 1 10 2zm0 1.3a.7.7 0 1 0 0 1.4.7.7 0 0 0 0-1.4z"/>',
        lock: '<path class="elx-lo-lock" d="M7 9V7a3 3 0 0 1 6 0v2h1a1 1 0 0 1 1 1v6a1 1 0 0 1-1 1H6a1 1 0 0 1-1-1v-6a1 1 0 0 1 1-1zm1.5 0h3V7a1.5 1.5 0 0 0-3 0z"/>',
        close: '<path class="elx-lo-x" d="M5.3 4.2 10 8.9l4.7-4.7 1.1 1.1-4.7 4.7 4.7 4.7-1.1 1.1-4.7-4.7-4.7 4.7-1.1-1.1 4.7-4.7-4.7-4.7z"/>'
    };
    var svg = function (name, size) {
        return '<svg class="elx-lo-icon" viewBox="0 0 20 20" width="' + size + '" height="' + size + '" aria-hidden="true">' + ICONS[name] + '</svg>';
    };

    var WRAPS = [
        { key: 'inline', style: 'Inline', text: 'textInline' },
        { key: 'square', style: 'Square', text: 'textSquare' },
        { key: 'tight', style: 'Tight', text: 'textTight' },
        { key: 'through', style: 'Through', text: 'textThrough' },
        { key: 'topAndBottom', style: 'TopAndBottom', text: 'textTopBottom' },
        { key: 'behind', style: 'Behind', text: 'textBehind' },
        { key: 'inFront', style: 'InFront', text: 'textInFront' }
    ];

    var STRINGS = {
        textTitle: 'Layout Options',
        textInlineGroup: 'In Line with Text',
        textWrapGroup: 'With Text Wrapping',
        textInline: 'In Line with Text',
        textSquare: 'Square',
        textTight: 'Tight',
        textThrough: 'Through',
        textTopBottom: 'Top and Bottom',
        textBehind: 'Behind Text',
        textInFront: 'In Front of Text',
        textMove: 'Move with text',
        textFix: 'Fix position on page',
        textSeeMore: 'See more...',
        textClose: 'Close',
        tipAnchor: 'Object anchor',
        tipLockedAnchor: 'Object anchor (locked)'
    };

    DE.Views.ElixirLayoutOptions = Backbone.View.extend(_.extend({

        initialize: function (options) {
            this.controller = options.controller;
            this.layout = null;
            this.key = '';
            this.selected = false;
        },

        setup: function (api, mode) {
            if (this.api || !api || !api.asc_GetSelectedDrawingLayout) return;
            this.api = api;
            this.mode = mode || {};
            var me = this;
            this.$main = $('#id_main_view');
            if (!this.$main.length) return;
            this.$btn = $('<button type="button" class="elx-lo-btn canfocused" aria-haspopup="dialog" aria-expanded="false"></button>')
                .attr('title', this.textTitle).attr('aria-label', this.textTitle).html(svg('square', 20)).hide();
            this.$anchor = $('<div class="elx-lo-anchor-mark" role="img"></div>').hide();
            this.$pane = $(this.renderPane()).hide();
            this.$main.append(this.$anchor, this.$btn, this.$pane);

            this.$btn.on('mousedown', function (e) { e.preventDefault(); e.stopPropagation(); });
            this.$btn.on('click', function (e) { e.preventDefault(); e.stopPropagation(); me.togglePane(); });
            this.$pane.on('mousedown', function (e) { e.stopPropagation(); if (!$(e.target).is('input')) e.preventDefault(); });
            this.$pane.on('click', '.elx-lo-tile', function () { me.applyWrap($(this).data('wrap')); });
            this.$pane.on('change', 'input[name=elx-lo-move]', function () { me.applyMove(this.value === 'move'); });
            this.$pane.on('focusout', 'input[name=elx-lo-move]', function (e) {
                // The next control in this flyout owns focus, rather than the document.
                if (e.relatedTarget && me.$pane[0].contains(e.relatedTarget)) e.stopPropagation();
            });
            this.$pane.on('click', '.elx-lo-more', function (e) { e.preventDefault(); me.seeMore(); });
            this.$pane.on('click', '.elx-lo-close', function () { me.hidePane(true); });
            this.$pane.on('keydown', function (e) {
                if (e.keyCode === Common.UI.Keys.ESC) { e.preventDefault(); e.stopPropagation(); me.hidePane(true); }  // the picture stays selected
                else if (e.keyCode === Common.UI.Keys.LEFT || e.keyCode === Common.UI.Keys.RIGHT) {
                    var $t = me.$pane.find('.elx-lo-tile'), i = $t.index(document.activeElement);
                    if (i >= 0) { e.preventDefault(); $t.eq((i + (e.keyCode === Common.UI.Keys.RIGHT ? 1 : $t.length - 1)) % $t.length).focus(); }
                }
            });
            // Esc closes the flyout wherever the focus is (the document keeps the picture selected)
            document.addEventListener('keydown', function (e) {
                if (e.keyCode === Common.UI.Keys.ESC && me.$pane.is(':visible')) {
                    e.preventDefault(); e.stopPropagation(); me.hidePane(true);
                }
            }, true);
            $(document).on('mousedown.elxlo', function (e) {
                if (me.$pane.is(':visible') && !$(e.target).closest('.elx-lo-pane, .elx-lo-btn').length) me.hidePane();
            });

            api.asc_registerCallback('asc_onFocusObject', _.bind(this.onFocusObject, this));
            api.asc_registerCallback('asc_onCoAuthoringDisconnect', function () { me.selected = false; me.refresh(); });
            // follows scrolling, zoom, recalculation and drags while an object is selected
            this.timer = setInterval(function () { me.selected && me.refresh(); }, 120);
        },

        renderPane: function () {
            var me = this;
            var tile = function (w) {
                return '<button type="button" class="elx-lo-tile" data-wrap="' + w.key + '" title="' + esc(me[w.text]) + '" aria-label="' + esc(me[w.text]) + '" aria-pressed="false">' + svg(w.key, 32) + '</button>';
            };
            return '<div class="elx-lo-pane canfocused" role="dialog" aria-label="' + esc(this.textTitle) + '">' +
                '<div class="elx-lo-head"><span class="elx-lo-title">' + esc(this.textTitle) + '</span>' +
                '<button type="button" class="elx-lo-close" title="' + esc(this.textClose) + '" aria-label="' + esc(this.textClose) + '">' + svg('close', 16) + '</button></div>' +
                '<div class="elx-lo-group">' + esc(this.textInlineGroup) + '</div>' +
                '<div class="elx-lo-tiles">' + tile(WRAPS[0]) + '</div>' +
                '<div class="elx-lo-group">' + esc(this.textWrapGroup) + '</div>' +
                '<div class="elx-lo-tiles">' + WRAPS.slice(1).map(tile).join('') + '</div>' +
                '<div class="elx-lo-sep"></div>' +
                '<label class="elx-lo-radio"><input type="radio" name="elx-lo-move" value="move"><span>' + esc(this.textMove) + '</span></label>' +
                '<label class="elx-lo-radio"><input type="radio" name="elx-lo-move" value="fix"><span>' + esc(this.textFix) + '</span></label>' +
                '<a href="#" class="elx-lo-more">' + esc(this.textSeeMore) + '</a>' +
                '</div>';
        },

        onFocusObject: function (selected) {
            var img = null;
            for (var i = selected.length - 1; i >= 0; i--) {
                if (selected[i].get_ObjectType() === Asc.c_oAscTypeSelectElement.Image) { img = selected[i].get_ObjectValue(); break; }
            }
            this.imgProps = img;
            this.selected = !!img;
            this.refresh();
        },

        canEdit: function () {
            return this.mode && this.mode.isEdit !== false && !(this.api.isViewMode) && !(this.imgProps && this.imgProps.get_Locked && this.imgProps.get_Locked());
        },

        refresh: function () {
            if (!this.$btn) return;
            var L = this.selected && this.canEdit() ? this.api.asc_GetSelectedDrawingLayout() : null;
            var key = L ? [L.left, L.top, L.right, L.bottom, L.wrap, L.moveWithText, L.lockAnchor, L.tracking, L.anchor ? L.anchor.x + ',' + L.anchor.y : ''].join('|') : '';
            if (key === this.key) return;
            this.key = key;
            this.layout = L;
            if (!L || L.tracking) {
                // hidden while the object is dragged, resized or rotated, like Word
                this.$btn.hide(); this.$anchor.hide();
                if (!L) this.hidePane();
                else this.$pane.hide();
                return;
            }
            var viewer = document.getElementById('id_viewer'), main = this.$main[0];
            if (!viewer || !main) return;
            var vr = viewer.getBoundingClientRect(), mr = main.getBoundingClientRect();
            var ox = vr.left - mr.left, oy = vr.top - mr.top, W = main.clientWidth, H = main.clientHeight;
            var size = 24, gap = 8;
            var x = ox + L.right + gap, y = oy + L.top;
            if (x + size > W - 4) x = ox + L.right - size - 4;              // no room on the right: inside
            if (x < 4) x = 4;
            y = Math.max(4, Math.min(y, oy + L.bottom - size));                 // top scrolled out: stay at the top of the view
            var visible = oy + L.bottom > size && oy + L.top < H - 4;
            this.$btn.css({ left: Math.round(x), top: Math.round(y) }).toggle(visible);

            if (L.anchor && !L.inline) {
                var ax = ox + L.anchor.x - 22, ay = oy + L.anchor.y;
                this.$anchor.html(svg('anchor', 16) + (L.lockAnchor ? '<span class="elx-lo-lockbadge">' + svg('lock', 10) + '</span>' : ''))
                    .attr('title', L.lockAnchor ? this.tipLockedAnchor : this.tipAnchor).attr('aria-label', L.lockAnchor ? this.tipLockedAnchor : this.tipAnchor)
                    .css({ left: Math.round(Math.max(0, ax)), top: Math.round(ay) })
                    .toggle(ay > -16 && ay < H);
            } else this.$anchor.hide();

            if (this.$pane.is(':visible')) this.placePane();
            this.updatePane();
        },

        togglePane: function () {
            if (this.$pane.is(':visible')) this.hidePane(true);
            else this.showPane();
        },

        showPane: function () {
            if (!this.layout) return;
            this.updatePane();
            this.$pane.show();
            this.placePane();
            this.$btn.attr('aria-expanded', 'true').addClass('active');
            var $cur = this.$pane.find('.elx-lo-tile.active');
            ($cur.length ? $cur : this.$pane.find('.elx-lo-tile').first()).focus();
        },

        hidePane: function (focusDoc) {
            if (!this.$pane) return;
            this.$pane.hide();
            this.$btn.attr('aria-expanded', 'false').removeClass('active');
            focusDoc && this.api && this.api.asc_enableKeyEvents(true);
        },

        placePane: function () {
            var b = this.$btn.position(), W = this.$main[0].clientWidth, H = this.$main[0].clientHeight;
            var pw = this.$pane.outerWidth(), ph = this.$pane.outerHeight();
            var x = b.left + this.$btn.outerWidth() + 6;
            if (x + pw > W - 4) x = b.left - pw - 6;
            var y = Math.min(b.top, H - ph - 4);
            this.$pane.css({ left: Math.max(4, Math.round(x)), top: Math.max(4, Math.round(y)) });
        },

        updatePane: function () {
            var L = this.layout;
            if (!L || !this.$pane) return;
            var cur = _.find(WRAPS, function (w) { return Asc.c_oAscWrapStyle2[w.style] === L.wrap; });
            this.$pane.find('.elx-lo-tile').each(function () {
                var on = cur && $(this).data('wrap') === cur.key;
                $(this).toggleClass('active', !!on).attr('aria-pressed', on ? 'true' : 'false');
            });
            var canFlow = !this.imgProps || !this.imgProps.get_CanBeFlow || this.imgProps.get_CanBeFlow();
            this.$pane.find('.elx-lo-tile').not('[data-wrap=inline]').prop('disabled', !canFlow);
            var $r = this.$pane.find('input[name=elx-lo-move]');
            $r.prop('disabled', L.inline);
            $r.filter('[value=move]').prop('checked', !L.inline && L.moveWithText);
            $r.filter('[value=fix]').prop('checked', !L.inline && !L.moveWithText);
        },

        // the ribbon's Wrap Text menu does the work (it keeps the place of a picture that
        // starts floating)
        applyWrap: function (key) {
            var w = _.findWhere(WRAPS, { key: key });
            var tb = DE.getController('Toolbar'), menu = tb && tb.toolbar && tb.toolbar.btnImgWrapping && tb.toolbar.btnImgWrapping.menu;
            if (!w || !menu) return;
            var item = _.find(menu.items, function (i) { return i.options.wrapType === Asc.c_oAscWrapStyle2[w.style]; });
            if (!item) return;
            menu.trigger('item:click', menu, item, {preserveFocus: true});
            this.key = '';
            this.refresh();
            this.$pane.find('.elx-lo-tile.active').focus();
        },

        // Word: "Fix position on page" ties the object to the page, "Move with text" to its
        // paragraph; the object stays where it is (values converted from the current place)
        applyMove: function (move) {
            var props = this.imgProps;
            if (!props || !this.api) return;
            var p = new Asc.asc_CImgProperty(), h = new Asc.CImagePositionH(), v = new Asc.CImagePositionV();
            var relH = move ? Asc.c_oAscRelativeFromH.Column : Asc.c_oAscRelativeFromH.Page,
                relV = move ? Asc.c_oAscRelativeFromV.Paragraph : Asc.c_oAscRelativeFromV.Page;
            h.put_UseAlign(false); h.put_Percent(false); h.put_RelativeFrom(relH); h.put_Value(props.get_Value_X(relH));
            v.put_UseAlign(false); v.put_Percent(false); v.put_RelativeFrom(relV); v.put_Value(props.get_Value_Y(relV));
            p.put_PositionH(h); p.put_PositionV(v);
            this.api.ImgApply(p);
            this.key = '';
            Common.NotificationCenter.trigger('edit:complete', this, {preserveFocus: true});
        },

        seeMore: function () {
            this.hidePane();
            var ct = DE.getController('ContextTabs');
            if (ct && ct.openImageAdvanced) ct.openImageAdvanced(9);
        }

    }, STRINGS, DE.Views.ElixirLayoutOptions || {}));

    DE.Views.ElixirLayoutOptions.ICONS = ICONS;
    return DE.Views.ElixirLayoutOptions;
});

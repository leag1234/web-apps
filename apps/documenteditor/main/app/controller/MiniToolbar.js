/*
 * (c) Copyright Ascensio System SIA 2010-2024
 *
 * This program is a free software product. You can redistribute it and/or
 * modify it under the terms of the GNU Affero General Public License (AGPL)
 * version 3 as published by the Free Software Foundation. In accordance with
 * Section 7(a) of the GNU AGPL its Section 15 shall be amended to the effect
 * that Ascensio System SIA expressly excludes the warranty of non-infringement
 * of any third-party rights.
 *
 * This program is distributed WITHOUT ANY WARRANTY; without even the implied
 * warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR  PURPOSE. For
 * details, see the GNU AGPL at: http://www.gnu.org/licenses/agpl-3.0.html
 *
 * The  interactive user interfaces in modified source and object code versions
 * of the Program must display Appropriate Legal Notices, as required under
 * Section 5 of the GNU AGPL version 3.
 *
 * All the Product's GUI elements, including illustrations and icon sets, as
 * well as technical writing content are licensed under the terms of the
 * Creative Commons Attribution-ShareAlike 4.0 International. See the License
 * terms at http://creativecommons.org/licenses/by-sa/4.0/legalcode
 *
 */

/**
 *  MiniToolbar.js
 *
 *  Shows Word's mini toolbar when text is selected with the mouse, near the
 *  pointer: it fades in, becomes opaque when the pointer comes close and
 *  fades out as the pointer moves away. It is also shown above the context
 *  menu of text and tables, as in Word.
 */

define([
    'core',
    'documenteditor/main/app/view/MiniToolbar'
], function () {
    'use strict';

    var FADE_DISTANCE = 150;   // px: the toolbar disappears beyond this distance from the pointer
    var IDLE_OPACITY = 0.55;   // opacity right after the selection, before the pointer comes close

    DE.Controllers.MiniToolbar = Backbone.Controller.extend({
        initialize: function() {
            Common.NotificationCenter.on('app:ready', _.bind(this.onAppReady, this));
        },

        onAppReady: function(mode) {
            this.mode = mode;
            if (!mode || !mode.isEdit) return;
            var me = this;
            this.api = this.getApplication().getController('Viewport').getApi();
            this.tbc = this.getApplication().getController('Toolbar');
            this.view = new DE.Views.MiniToolbar({toolbar: this.tbc.toolbar}).render();
            this.wire();

            var $sdk = $('#editor_sdk');
            // the editor captures the pointer while selecting (and may cancel the mouse events):
            // listen to the pointer events on the window, capture phase
            var downInDocument = false;
            window.addEventListener('pointerdown', function(e) {
                downInDocument = e.button === 0 && !!$(e.target).closest('#editor_sdk').length;
            }, true);
            window.addEventListener('pointerup', function(e) {
                if (e.button !== 0 || !downInDocument) return;
                downInDocument = false;
                var x = e.pageX, y = e.pageY;
                // after the editor has updated the selection
                setTimeout(function() { me.onSelectionMouseUp(x, y); }, 30);
            }, true);
            window.addEventListener('pointermove', function(e) { me.onMouseMove(e.pageX, e.pageY); }, true);
            $(document).on('mousedown', function(e) {
                if (me.view.$el[0].contains(e.target) || $(e.target).closest('.dropdown-menu').length && me.isOwnMenu(e.target)) {
                    me.interacting = true;
                    return;
                }
                me.interacting = false;
                me.hide();
            });
            $(document).on('keydown', function(e) {
                if (me.view.$el[0].contains(e.target) || me.isOwnMenu(e.target)) return;
                me.hide();
            });
            $sdk.on('wheel', function() { me.hide(); });
            Common.NotificationCenter.on('window:resize', function() { me.hide(); });
            this.api.asc_registerCallback('asc_onContextMenu', _.bind(this.onContextMenu, this));
            // the font of the selection arrives from the engine (after it has loaded the font): the ribbon
            // combos update first (own timer of 100 ms), then the mini toolbar shows the same
            var timer = null, later = function() {
                if (!me.isVisible()) return;
                clearTimeout(timer);
                timer = setTimeout(function() { me.isVisible() && me.syncState(); }, 160);
            };
            Common.NotificationCenter.on('fonts:change', later);
            this.api.asc_registerCallback('asc_onFontSize', later);
        },

        isOwnMenu: function(el) {
            var vw = this.view, menus = [vw.btnStyles.menu, vw.btnHighlight.menu, vw.btnFontColor.menu,
                vw.cmbFontName, vw.cmbFontSize, vw.btnBorders.menu, vw.btnShading.menu, vw.btnInsert.menu, vw.btnDelete.menu];
            return _.some(menus, function(m) { return m && m.cmpEl && m.cmpEl[0] && m.cmpEl[0].contains(el); });
        },

        wire: function() {
            var me = this, v = this.view, api = this.api, tbc = this.tbc;
            var done = function() { me.syncState(); };
            v.cmbFontName.on('selected', function(combo, record) { tbc.onFontNameSelect(combo, record); done(); });
            v.cmbFontSize.on('selected', function(combo, record) { tbc.onFontSizeSelect(combo, record); done(); });
            v.cmbFontSize.on('changed:after', function(combo, record) {
                var size = Common.Utils.String.parseFloat(record.value);
                size > 0 && api.put_TextPrFontSize(Math.min(size, 300));
                done();
            });
            v.btnGrow.on('click', function() { api.FontSizeIn(); done(); });
            v.btnShrink.on('click', function() { api.FontSizeOut(); done(); });
            v.btnPainter.on('click', function(btn) {
                tbc.toolbar.btnCopyStyle.toggle(btn.pressed, true);
                tbc.onCopyStyleToggle(btn, btn.pressed);
                me.hide(true);
            });
            v.btnBold.on('click', function(btn) { api.put_TextPrBold(btn.pressed); done(); });
            v.btnItalic.on('click', function(btn) { api.put_TextPrItalic(btn.pressed); done(); });
            v.btnUnderline.on('click', function(btn) { api.put_TextPrUnderline(btn.pressed); done(); });
            v.btnHighlight.on('click', function() { tbc._setMarkerColor(tbc.toolbar.btnHighlightColor.currentColor || 'FFFF00'); done(); });
            v.on('highlight:select', function(color) {
                tbc.toolbar.btnHighlightColor.currentColor = color;
                tbc.toolbar.btnHighlightColor.setColor(color);
                v.btnHighlight.setColor(color);
                tbc._setMarkerColor(color, 'menu');
                done();
            });
            v.mnuHighlightNone.on('click', function() { api.SetMarkerFormat(true, false); done(); });
            v.btnFontColor.on('click', function() { tbc.onBtnFontColor(); done(); });
            v.btnFontColor.on('color:select', function(btn, color) { tbc.onSelectFontColor(btn, color); v.btnFontColor.setColor(color.color || color); done(); });
            v.btnFontColor.on('auto:select', function(btn, color) { tbc.onAutoFontColor && tbc.onAutoFontColor(btn, color); done(); });
            v.btnBullets.on('click', function(btn) { tbc.onMarkers(btn); done(); });
            v.btnNumbering.on('click', function(btn) { tbc.onNumbers(btn); done(); });
            // Elixir AI: hidden when the AI layer is not available in this editor
            var ai = me.getApplication().getController('Common.Controllers.ElixirAI');
            var aiShow = function() { v.btnElixir.setVisible(!!(ai && ai.isActive && ai.isActive())); };
            aiShow();
            Common.NotificationCenter.on('elixir:ai-state', aiShow);
            v.btnElixir.on('click', function() {
                var ep = Common.ElixirAI && Common.ElixirAI.entryPoints && Common.ElixirAI.entryPoints.word;
                var off = v.btnElixir.cmpEl.offset(), h = v.btnElixir.cmpEl.outerHeight();
                me.hide(true);
                ep && ep.openAskMenu && ep.openAskMenu(off.left, off.top + h + 4);
            });
            var comment = function() {
                me.hide(true);
                var dh = me.getApplication().getController('DocumentHolder');
                dh && dh.addComment && dh.addComment();
            };
            v.btnComment.on('click', comment);
            v.btnTableComment.on('click', comment);
            // table mode: borders, shading, Insert, Delete (same commands as the Table Design and Layout tabs)
            var ctx = function() { return me.getApplication().getController('ContextTabs'); };
            var applyBorders = function(code) { var c = ctx(); c && c.applyCellBorders(code); done(); };
            v.btnBorders.menu.on('item:click', function(menu, item) {
                if (item.value === 'advanced') { me.hide(true); var c = ctx(); c && c.openTableAdvanced(2); return; }
                v.btnBorders.options.lastBorder = item.value;
                applyBorders(item.value);
            });
            v.btnBorders.on('click', function() { applyBorders(v.btnBorders.options.lastBorder); });
            var shade = function(color) {
                var p = new Asc.CTableProp(), bg = new Asc.CBackground();
                if (color == 'transparent') bg.put_Value(1);
                else { bg.put_Value(0); bg.put_Color(Common.Utils.ThemeColor.getRgbColor(color)); }
                p.put_CellsBackground(bg);
                p.put_CellSelect(true);
                api.tblApply(p);
                done();
            };
            v.btnShading.on('color:select', function(btn, color) { shade(color); });
            v.btnShading.on('click', function() { shade(v.btnShading.color || 'transparent'); });
            v.btnInsert.menu.on('item:click', function(menu, item) {
                ({above: api.addRowAbove, below: api.addRowBelow, left: api.addColumnLeft, right: api.addColumnRight})[item.value].call(api);
                me.hide(true);
            });
            v.btnDelete.menu.on('item:click', function(menu, item) {
                me.hide(true);
                if (item.value === 'cells') { var c = ctx(); c && c.dh().onCellsRemove(); }
                else ({column: api.remColumn, row: api.remRow, table: api.remTable})[item.value].call(api);
            });
            v.btnStyles.menu.on('show:before', function(menu) { me.fillStyles(menu); });
            // (on a part of a paragraph, the linked character style is applied, as in Word)
            v.btnStyles.menu.on('item:click', function(menu, item) {
                api.put_Style(item.value);
                done();
            });
        },

        fillStyles: function(menu) {
            menu.removeAll();
            var list = this.tbc.toolbar.listStyles, store = list && list.menuPicker ? list.menuPicker.store : null;
            var current = list && list.menuPicker && list.menuPicker.getSelectedRec ? list.menuPicker.getSelectedRec() : null;
            store && store.each(function(rec) {
                // the gallery records: title = style name, tip = displayed (translated) name
                var name = rec.get('title'), caption = rec.get('tip') || name;
                menu.addItem(new Common.UI.MenuItem({caption: caption, value: name, checkable: true, toggleGroup: 'mini-styles', checked: current === rec}));
            });
        },

        // copy the toolbar's state (kept up to date by the toolbar controller)
        syncState: function() {
            var tb = this.tbc.toolbar, v = this.view;
            try {
                // the ribbon fills its size list after the mini toolbar was created: copy it (an empty list opens as a thin empty box)
                var sizes = tb.cmbFontSize.store;
                if (sizes.length && v.cmbFontSize.store.length !== sizes.length && !v.cmbFontSize.isMenuOpen())
                    v.cmbFontSize.setData(sizes.toJSON());
                // what the ribbon shows (its getValue() keeps the last record picked in the ribbon, not the font of the selection)
                v.cmbFontName.setValue(tb.cmbFontName.getRawValue());
                v.cmbFontSize.setValue(tb.cmbFontSize.getRawValue());
                v.btnBold.toggle(!!tb.btnBold.pressed, true);
                v.btnItalic.toggle(!!tb.btnItalic.pressed, true);
                v.btnUnderline.toggle(!!tb.btnUnderline.pressed, true);
                v.btnBullets.toggle(!!tb.btnMarkers.pressed, true);
                v.btnNumbering.toggle(!!tb.btnNumbers.pressed, true);
                v.btnPainter.toggle(!!tb.btnCopyStyle.pressed, true);
                tb.btnHighlightColor.currentColor && v.btnHighlight.setColor(tb.btnHighlightColor.currentColor);
                var fc = tb.btnFontColor.currentColor;
                fc && v.btnFontColor.setColor(typeof fc === 'object' ? (fc.isAuto ? '000' : fc.color) : fc);
            } catch (e) {}
        },

        // the selection is text, not a picture or a shape
        isTextSelection: function() {
            if (!this.api || !this.api.can_CopyCut || !this.api.can_CopyCut()) return false;
            var objects = this.api.getSelectedElements ? this.api.getSelectedElements() : [];
            for (var i = 0; i < objects.length; i++) {
                var type = objects[i].get_ObjectType();
                if (type === Asc.c_oAscTypeSelectElement.Image || type === Asc.c_oAscTypeSelectElement.Chart) {
                    var value = objects[i].get_ObjectValue();
                    // inside a text box the text is selected: keep the toolbar
                    if (!(value && value.get_ShapeProperties && value.get_ShapeProperties() && this.api.asc_GetSelectedText && this.api.asc_GetSelectedText()))
                        return false;
                }
            }
            return true;
        },

        canShow: function() {
            return this.mode && this.mode.isEdit && !Common.Utils.ModalWindow.isVisible() && !$('#file-menu-panel').is(':visible') &&
                !this.getApplication().getController('Toolbar').toolbar.btnCopyStyle.pressed;
        },

        onSelectionMouseUp: function(x, y) {
            if (!this.canShow() || !this.isTextSelection()) {
                this.hide();
                return;
            }
            this.showAt(x, y, false);
        },

        onContextMenu: function(event) {
            var me = this;
            // after DocumentHolder has opened its menu
            setTimeout(function() {
                if (!me.canShow()) return;
                var dh = me.getApplication().getController('DocumentHolder');
                var view = dh && dh.documentHolder, menu = view && view.currentMenu;
                if (!menu || !menu.isVisible() || (menu !== view.textMenu && menu !== view.tableMenu)) return;
                var $menu = menu.cmpEl.find('> .dropdown-menu').addBack('.dropdown-menu').first();
                if (!$menu.length) $menu = menu.cmpEl;
                var off = $menu.offset();
                me.view.setTableMode(menu === view.tableMenu);
                me.showAbove(off.left, off.top, $menu.outerWidth());
                menu.once('hide:after', function() { _.defer(function() { if (!me.interacting) me.hide(true); }); });
            }, 20);
        },

        place: function(left, top) {
            var $el = this.view.$el, w = $el.outerWidth(), h = $el.outerHeight();
            var maxLeft = Common.Utils.innerWidth() - w - 4, maxTop = Common.Utils.innerHeight() - h - 4;
            $el.css({left: Math.max(4, Math.min(left, maxLeft)), top: Math.max(4, Math.min(top, maxTop))});
        },

        showAt: function(x, y, opaque) {
            this.view.setTableMode(false);
            this.syncState();
            var $el = this.view.$el.addClass('visible');
            var h = $el.outerHeight();
            // above the pointer, slightly to the left as in Word; below if there is no room
            var top = y - h - 18;
            if (top < $('#editor_sdk').offset().top + 4) top = y + 22;
            this.place(x - 24, top);
            this.fromMenu = false;
            this.setOpacity(opaque ? 1 : IDLE_OPACITY);
            this.onMouseMove(x, y);
        },

        showAbove: function(left, menuTop) {
            this.syncState();
            var $el = this.view.$el.addClass('visible');
            this.place(left, menuTop - $el.outerHeight() - 6);
            this.fromMenu = true;
            this.setOpacity(1);
        },

        setOpacity: function(value) {
            this.view.$el.css('opacity', value);
        },

        onMouseMove: function(x, y) {
            var $el = this.view && this.view.$el;
            if (!$el || !$el.hasClass('visible') || this.fromMenu) return;
            var off = $el.offset(), w = $el.outerWidth(), h = $el.outerHeight();
            var dx = Math.max(off.left - x, 0, x - (off.left + w)), dy = Math.max(off.top - y, 0, y - (off.top + h));
            var d = Math.sqrt(dx * dx + dy * dy);
            if (d === 0 || this.isMenuOpen()) {
                this.setOpacity(1);
            } else if (d > FADE_DISTANCE) {
                this.hide();
            } else {
                // closer is more opaque; fully opaque only under the pointer
                var k = 1 - d / FADE_DISTANCE;
                this.setOpacity(Math.max(0.1, this.wasNear ? Math.min(0.95, k) : IDLE_OPACITY * (0.5 + k / 2)));
            }
            if (d < 30) this.wasNear = true;
        },

        isMenuOpen: function() {
            var v = this.view;
            return v.btnStyles.menu.isVisible() || (v.btnHighlight.menu && v.btnHighlight.menu.isVisible && v.btnHighlight.menu.isVisible()) ||
                (v.btnFontColor.menu && v.btnFontColor.menu.isVisible && v.btnFontColor.menu.isVisible()) ||
                v.cmbFontName.isMenuOpen() || v.cmbFontSize.isMenuOpen() ||
                _.some([v.btnBorders, v.btnShading, v.btnInsert, v.btnDelete], function(b) { return b.menu && b.menu.isVisible && b.menu.isVisible(); });
        },

        hide: function(force) {
            var $el = this.view && this.view.$el;
            if (!$el || !$el.hasClass('visible')) return;
            if (!force && this.isMenuOpen()) return;
            $el.removeClass('visible');
            this.wasNear = false;
            this.fromMenu = false;
        },

        isVisible: function() {
            return !!(this.view && this.view.$el.hasClass('visible'));
        }
    });
});

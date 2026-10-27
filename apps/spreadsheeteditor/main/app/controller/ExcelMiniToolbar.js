/*
 * SPDX-FileCopyrightText: 2026 ElixirOffice contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 *
 * SSE.Controllers.ExcelMiniToolbar: Excel's mini toolbar, shown above the
 * context menu of the cells (right-click), like Excel:
 *   Font, Size, Grow Font, Shrink Font, Accounting, Percent, Comma, Format Painter
 *   Bold, Italic, Center, Fill Color, Font Color, Borders, Increase / Decrease
 *   Decimal, Merge & Center.
 * Every command goes through the ribbon control it mirrors (same handlers,
 * locks and undo), so the result is the ribbon's. Styled like the document
 * editor's mini toolbar (#mini-toolbar, theme interact.less).
 */

define([
    'core',
    'common/main/lib/component/Button',
    'common/main/lib/component/ComboBoxFonts'
], function () {
    'use strict';

    SSE.Controllers.ExcelMiniToolbar = Backbone.Controller.extend(_.extend({
        initialize: function () {
            Common.NotificationCenter.on('app:ready', _.bind(this.onAppReady, this));
        },

        onAppReady: function (mode) {
            this.mode = mode;
            if (!mode || !mode.isEdit || mode.isEditDiagram || mode.isEditMailMerge || mode.isEditOle) return;
            var me = this;
            this.api = this.getApplication().getController('Viewport').getApi();
            this.tbc = this.getApplication().getController('Toolbar');
            this.tb = this.tbc && this.tbc.toolbar;
            if (!this.api || !this.tb) return;
            this.render();
            $(document).on('mousedown', function (e) {
                if (me.$el && (me.$el[0].contains(e.target) || me.isOwnMenu(e.target))) { me.interacting = true; return; }
                me.interacting = false;
                me.hide(true);
            });
            $(document).on('keydown', function (e) {
                if (me.$el && (me.$el[0].contains(e.target) || me.isOwnMenu(e.target))) return;
                me.hide(true);
            });
            Common.NotificationCenter.on('window:resize', function () { me.hide(true); });
            this.api.asc_registerCallback('asc_onContextMenu', _.bind(this.onContextMenu, this));
        },

        render: function () {
            var me = this, tb = this.tb, T = SSE.Views.Toolbar.prototype;
            var $el = $('<div id="mini-toolbar" class="mini-toolbar elx-excel-mini" role="toolbar" aria-label="' + Common.Utils.String.htmlEncode(this.textMiniToolbar) + '">' +
                '<div class="mini-row"><span class="mini-slot" id="mini-font-name"></span><span class="mini-slot" id="mini-font-size"></span>' +
                '<span class="mini-slot" id="mini-grow"></span><span class="mini-slot" id="mini-shrink"></span>' +
                '<span class="mini-slot" id="mini-accounting"></span><span class="mini-slot" id="mini-percent"></span>' +
                '<span class="mini-slot" id="mini-comma"></span><span class="mini-slot" id="mini-painter"></span></div>' +
                '<div class="mini-row"><span class="mini-slot" id="mini-bold"></span><span class="mini-slot" id="mini-italic"></span>' +
                '<span class="mini-slot" id="mini-center"></span><span class="mini-slot" id="mini-fill"></span>' +
                '<span class="mini-slot" id="mini-fontcolor"></span><span class="mini-slot" id="mini-borders"></span>' +
                '<span class="mini-slot" id="mini-incdec"></span><span class="mini-slot" id="mini-decdec"></span>' +
                '<span class="mini-slot" id="mini-merge"></span></div>' +
                '</div>');
            $(document.body).append($el);
            this.$el = $el;

            var ai=this.getApplication().getController('Common.Controllers.ElixirAI');
            if(ai && ai.isAvailable()) {
                var slot=$('<span class="mini-slot" id="mini-elixir-ask"></span>');
                $el.find('.mini-row').last().append(slot);
                this.btnElixirAsk=new Common.UI.Button({parentEl:slot,cls:'btn-toolbar',iconCls:'elixir-ai-icon elixir-ai-icon-sparkle',hint:Common.ElixirAI.t('menuAsk')});
                this.btnElixirAsk.on('click',function(){me.hide(true);ai.openBar();});
            }

            // a button forwarding its click to a ribbon control
            var twin = function (slot, src, opts) {
                if (!src) return null;
                var b = new Common.UI.Button(_.extend({
                    parentEl: $el.find('#' + slot),
                    cls: 'btn-toolbar',
                    iconCls: src.options.iconCls,
                    hint: src.options.hint,
                    enableToggle: !!src.options.enableToggle
                }, opts || {}));
                b.on('click', function () {
                    if (src.isDisabled()) return;
                    if (src.options.enableToggle) src.toggle(b.pressed, true);
                    src.trigger('click', src, {});
                    me.afterCommand();
                });
                b._src = src;
                return b;
            };

            this.cmbFontName = new Common.UI.ComboBoxFonts({
                el: $el.find('#mini-font-name'),
                cls: 'input-group-nr',
                menuCls: 'scrollable-menu',
                menuStyle: 'min-width: 325px;',
                hint: T.tipFontName,
                store: tb.cmbFontName ? tb.cmbFontName.store : new Common.Collections.Fonts(),
                recent: 0
            });
            this.cmbFontSize = new Common.UI.ComboBox({
                el: $el.find('#mini-font-size'),
                cls: 'input-group-nr',
                menuCls: 'scrollable-menu',
                menuStyle: 'min-width: 55px;max-height: 300px;',
                hint: T.tipFontSize,
                data: tb.cmbFontSize ? tb.cmbFontSize.store.toJSON() : []
            });
            this.cmbFontName.on('selected', function (combo, record) { me.tbc.onFontNameSelect(combo, record); me.afterCommand(); });
            this.cmbFontSize.on('selected', function (combo, record) { me.tbc.onFontSizeSelect(combo, record); me.afterCommand(); });
            this.cmbFontSize.on('changed:after', function (combo, record) {
                var size = Common.Utils.String.parseFloat(record.value);
                size > 0 && me.api.asc_setCellFontSize(Math.min(size, 409));
                me.afterCommand();
            });

            this.buttons = _.compact([
                twin('mini-grow', tb.btnIncFontSize), twin('mini-shrink', tb.btnDecFontSize),
                twin('mini-accounting', tb.btnCurrencyStyle), twin('mini-percent', tb.btnPercentStyle), twin('mini-comma', tb.btnCommaStyle),
                this.btnPainter = twin('mini-painter', tb.btnCopyStyle, {enableToggle: true}),
                this.btnBold = twin('mini-bold', tb.btnBold), this.btnItalic = twin('mini-italic', tb.btnItalic),
                twin('mini-center', tb.btnAlignCenter, {enableToggle: false}),
                twin('mini-borders', tb.btnBorders), twin('mini-incdec', tb.btnIncDecimal), twin('mini-decdec', tb.btnDecDecimal),
                twin('mini-merge', tb.btnMerge, {enableToggle: false})
            ]);
            // the format painter stays on until the next paste, like the ribbon's
            this.btnPainter && this.btnPainter.off('click') && this.btnPainter.on('click', function (b) {
                var src = tb.btnCopyStyle;
                if (!src || src.isDisabled()) return;
                src.toggle(b.pressed);
                me.hide(true);
            });

            var colored = function (slot, src, onSelect) {
                if (!src) return null;
                var b = new Common.UI.ButtonColored({
                    parentEl: $el.find('#' + slot),
                    cls: 'btn-toolbar',
                    iconCls: src.options.iconCls,
                    hint: src.options.hint,
                    split: true,
                    menu: true,
                    transparent: slot === 'mini-fill',
                    auto: slot === 'mini-fontcolor'
                });
                b.setMenu();
                var picker = b.getPicker();
                b.menu.on('show:after', function () {
                    picker.updateColors(Common.Utils.ThemeColor.getEffectColors(), Common.Utils.ThemeColor.getStandartColors());
                });
                b.on('click', function () {
                    var c = src.currentColor;
                    c !== undefined && onSelect(src, c);
                    me.afterCommand();
                });
                b.on('color:select', function (btn, color) {
                    src.currentColor = color;
                    src.setColor && src.setColor(color.color || color);
                    b.setColor(color.color || color);
                    onSelect(src, color);
                    me.afterCommand();
                });
                b.on('auto:select', function (btn, color) {
                    me.tbc.onAutoFontColor && me.tbc.onAutoFontColor(src, color);
                    me.afterCommand();
                });
                return b;
            };
            this.btnFill = colored('mini-fill', tb.btnBackColor, function (src, c) { me.tbc.onBackColorSelect(src, c); });
            this.btnFontColor = colored('mini-fontcolor', tb.btnTextColor, function (src, c) { me.tbc.onTextColorSelect(src, c); });
            return this;
        },

        isOwnMenu: function (el) {
            var menus = [this.btnFill && this.btnFill.menu, this.btnFontColor && this.btnFontColor.menu, this.cmbFontName, this.cmbFontSize];
            return _.some(menus, function (m) { return m && m.cmpEl && m.cmpEl[0] && m.cmpEl[0].contains(el); });
        },

        isMenuOpen: function () {
            return !!((this.btnFill && this.btnFill.menu && this.btnFill.menu.isVisible && this.btnFill.menu.isVisible()) ||
                (this.btnFontColor && this.btnFontColor.menu && this.btnFontColor.menu.isVisible && this.btnFontColor.menu.isVisible()) ||
                this.cmbFontName.isMenuOpen() || this.cmbFontSize.isMenuOpen());
        },

        afterCommand: function () {
            var me = this;
            // Excel closes the context menu and keeps the mini toolbar for more formatting
            Common.UI.Menu.Manager.hideAll();
            _.defer(function () { me.syncState(); });
        },

        syncState: function () {
            var tb = this.tb;
            try {
                this.cmbFontName.setValue(tb.cmbFontName.getValue());
                this.cmbFontSize.setValue(tb.cmbFontSize.getValue());
                _.each(this.buttons, function (b) {
                    b.setDisabled(!!b._src.isDisabled());
                    b.options.enableToggle && b.toggle(!!b._src.pressed, true);
                });
                tb.btnBackColor && tb.btnBackColor.currentColor && this.btnFill.setColor(tb.btnBackColor.currentColor.color || tb.btnBackColor.currentColor);
                var fc = tb.btnTextColor && tb.btnTextColor.currentColor;
                fc && this.btnFontColor.setColor(typeof fc === 'object' ? (fc.isAuto ? '000' : fc.color) : fc);
            } catch (e) {}
        },

        canShow: function () {
            return this.mode && this.mode.isEdit && !Common.Utils.ModalWindow.isVisible() && !$('#file-menu-panel').is(':visible') &&
                !(this.tb.btnCopyStyle && this.tb.btnCopyStyle.pressed);
        },

        onContextMenu: function () {
            var me = this;
            // after DocumentHolder has opened its menu
            setTimeout(function () {
                if (!me.canShow()) return;
                var dh = me.getApplication().getController('DocumentHolder'),
                    view = dh && dh.documentHolder, menu = dh && dh.currentMenu;
                if (!menu || !view || menu !== view.ssMenu || !menu.isVisible()) return;
                var $menu = menu.cmpEl.find('> .dropdown-menu').addBack('.dropdown-menu').first();
                if (!$menu.length) $menu = menu.cmpEl;
                var off = $menu.offset();
                me.syncState();
                me.$el.addClass('visible').css('opacity', 1);
                var w = me.$el.outerWidth(), h = me.$el.outerHeight(),
                    top = off.top - h - 6;
                // no room above: under the menu
                if (top < 4) top = off.top + $menu.outerHeight() + 6;
                me.$el.css({left: Math.max(4, Math.min(off.left, Common.Utils.innerWidth() - w - 4)), top: Math.max(4, top)});
            }, 30);
        },

        hide: function (force) {
            if (!this.$el || !this.$el.hasClass('visible')) return;
            if (!force && this.isMenuOpen()) return;
            if (this.interacting && !force) return;
            this.$el.removeClass('visible');
        },

        textMiniToolbar: 'Mini toolbar'
    }, SSE.Controllers.ExcelMiniToolbar || {}));
});

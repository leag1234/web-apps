/*
 * SPDX-FileCopyrightText: 2026 ElixirOffice contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 *
 * PE.Controllers.SlideDesignGalleries: Design > Variants gets Fonts and Background
 * Styles next to Colors, like PowerPoint's Variants menu.
 *  - Fonts: the theme font schemes of the Office list (heading font / body font); applied to
 *    every theme of the presentation through api.asc_ChangeFontScheme (sdkjs slide/api.js).
 *  - Background Styles: 12 styles (solid, subtle and intense gradients of the theme colours),
 *    set on the current slide and applied to all slides, as the Format Background pane does.
 */
define([
    'core',
    'common/main/lib/component/Button',
    'common/main/lib/component/DataView'
], function () {
    'use strict';

    // Office's "Fonts" list: [name, heading font, body font]
    var FONT_SCHEMES = [
        ['Office', 'Calibri Light', 'Calibri'],
        ['Office 2007 - 2010', 'Calibri', 'Cambria'],
        ['Calibri', 'Calibri', 'Calibri'],
        ['Arial', 'Arial', 'Arial'],
        ['Corbel', 'Corbel', 'Corbel'],
        ['Candara', 'Candara', 'Candara'],
        ['Franklin Gothic', 'Franklin Gothic Medium', 'Franklin Gothic Book'],
        ['Century Gothic', 'Century Gothic', 'Century Gothic'],
        ['Tw Cen MT', 'Tw Cen MT', 'Tw Cen MT'],
        ['Cambria', 'Cambria', 'Calibri'],
        ['Garamond', 'Garamond', 'Garamond'],
        ['Georgia', 'Georgia', 'Georgia'],
        ['Gill Sans MT', 'Gill Sans MT', 'Gill Sans MT'],
        ['Trebuchet MS', 'Trebuchet MS', 'Trebuchet MS'],
        ['Verdana', 'Verdana', 'Verdana'],
        ['Arial Black - Arial', 'Arial Black', 'Arial'],
        ['Consolas - Corbel', 'Consolas', 'Corbel'],
        ['Constantia - Franklin Gothic Book', 'Constantia', 'Franklin Gothic Book'],
        ['Corbel - Cambria', 'Corbel', 'Cambria'],
        ['Times New Roman - Arial', 'Times New Roman', 'Arial']
    ];

    // Background styles: [first colour, second colour or null], colours are indexes of the first row
    // of the theme colour palette: 0 Background 1, 1 Text 1, 2 Background 2, 3 Text 2, 4 Accent 1, 5 Accent 2
    var BG_STYLES = [
        [[0], [2], [3], [1]],
        [[0, 2], [2, 0], [3, 1], [1, 3]],
        [[0, 4], [2, 5], [3, 4], [1, 5]]
    ];

    var STRINGS = {
        txtFonts: 'Fonts', txtBackgroundStyles: 'Background Styles', txtStyle: 'Style',
        tipFonts: 'Change the fonts of the theme: headings and body text.',
        tipBackgroundStyles: 'Pick a background style for all the slides.'
    };

    PE.Views.SlideDesignGalleries = Backbone.View.extend(_.extend({ initialize: function () {} }, STRINGS, PE.Views.SlideDesignGalleries || {}));
    PE.Views.SlideDesignGalleries.STRINGS = STRINGS;
    PE.Views.SlideDesignGalleries.FONT_SCHEMES = FONT_SCHEMES;
    PE.Views.SlideDesignGalleries.BG_STYLES = BG_STYLES;

    PE.Controllers.SlideDesignGalleries = Backbone.Controller.extend({

        initialize: function () {
            var me = this;
            this.addListeners({
                'Toolbar': { 'render:after': function (toolbar) { me.onToolbarRender(toolbar); } }
            });
            Common.NotificationCenter.on('app:ready', this.onAppReady.bind(this));
        },

        onToolbarRender: function (tb) {
            var main = this.getApplication().getController('Main'), mode = (main && main.appOptions) || {};
            if (!mode.isEdit || this.tb) return;
            var $colors = tb.$el.find('#slot-btn-colorschemas');
            if (!$colors.length) return;
            this.tb = tb;
            this.strings = new PE.Views.SlideDesignGalleries();
            this.controls = [];
            var me = this, s = this.strings, _set = Common.enumLock;
            var big = function (id, cfg) {
                var $slot = $('<span class="btn-slot text x-huge" id="' + id + '"></span>');
                $colors.after($slot);
                var b = new Common.UI.Button(_.extend({
                    cls: 'btn-toolbar x-huge icon-top', menu: true,
                    lock: [_set.lostConnect, _set.noSlides, _set.disableOnStart],
                    dataHint: '1', dataHintDirection: 'bottom', dataHintOffset: 'small'
                }, cfg));
                b.render($slot);
                me.controls.push(b);
                return b;
            };
            // after Colors: Colors, Fonts, Background Styles
            this.btnBg = big('slot-elx-bgstyles', { iconCls: 'toolbar__icon btn-elx-bg-styles', caption: s.txtBackgroundStyles, hint: s.tipBackgroundStyles });
            this.btnFonts = big('slot-elx-fonts', { iconCls: 'toolbar__icon btn-elx-theme-fonts', caption: s.txtFonts, hint: s.tipFonts });

            this.btnFonts.setMenu(new Common.UI.Menu({
                style: 'max-height: 380px; overflow-y: auto;',
                items: _.map(FONT_SCHEMES, function (f) {
                    return { caption: f[0], value: f, checkable: true, toggleGroup: 'elx-theme-fonts', style: "font-family: '" + f[1] + "';" };
                })
            }));
            this.btnFonts.menu.on('show:before', function (menu) { me.markCurrentFonts(menu); });
            this.btnFonts.menu.on('item:click', function (menu, item) { me.applyFonts(item.value); });

            this.btnBg.setMenu(new Common.UI.Menu({
                style: 'min-width: 0;',
                items: [{ template: _.template('<div id="elx-bgstyles-view" style="width: 200px; margin: 0 4px;"></div>') }]
            }));
            this.btnBg.menu.on('show:before', function (menu) { me.buildBackgrounds(menu); });

            Array.prototype.push.apply(tb.lockControls, this.controls);
            Common.Utils.lockControls(_set.disableOnStart, true, { array: this.controls });
        },

        onAppReady: function () {
            if (!this.tb) return;
            this.api = this.getApplication().getController('Viewport').getApi();
            Common.Utils.lockControls(Common.enumLock.disableOnStart, false, { array: this.controls });
            var me = this;
            this.api.asc_registerCallback('asc_onSendThemeColors', function () { me._bgBuilt && setTimeout(function () { me.updateBackgrounds(); }, 0); });
        },

        // ---- Fonts ---------------------------------------------------------------------------
        markCurrentFonts: function (menu) {
            var cur = this.api && this.api.asc_getThemeFonts && this.api.asc_getThemeFonts();
            _.each(menu.items, function (it) {
                var v = it.value;
                it.setChecked(!!(cur && v && v[1] === cur[0] && v[2] === cur[1]), true);
            });
        },

        applyFonts: function (scheme) {
            if (!this.api || !scheme || !this.api.asc_ChangeFontScheme) return;
            this.api.asc_ChangeFontScheme(scheme[1], scheme[2], scheme[0]);
            Common.NotificationCenter.trigger('edit:complete', this.tb);
        },

        // ---- Background Styles -----------------------------------------------------------------
        palette: function () { return Common.Utils.ThemeColor.getEffectColors() || []; },

        cssOf: function (spec) {
            var pal = this.palette(), hex = function (i) { return '#' + (pal[i] && pal[i].color || 'ffffff').replace('#', ''); };
            return spec.length === 1 ? hex(spec[0]) : 'linear-gradient(to bottom, ' + hex(spec[0]) + ', ' + hex(spec[1]) + ')';
        },

        buildBackgrounds: function (menu) {
            if (this._bgBuilt) return;
            var me = this, n = 0, models = [];
            _.each(BG_STYLES, function (row) {
                _.each(row, function (spec) {
                    models.push({ spec: spec, num: ++n, css: me.cssOf(spec) });
                });
            });
            this._bgStore = new Common.UI.DataViewStore(models);
            this._bgBuilt = true;
            this._bgPicker = new Common.UI.DataView({
                el: $('#elx-bgstyles-view', menu.$el),
                store: this._bgStore,
                parentMenu: menu,
                outerMenu: { menu: menu, index: 0 },
                showLast: false,
                itemTemplate: _.template('<div class="elx-bgstyle" title="' + this.strings.txtStyle + ' <%= num %>" style="width: 40px; height: 30px; background: <%= css %>; border: 1px solid #b5b5b5; box-sizing: border-box;"></div>')
            });
            this._bgPicker.on('item:click', function (picker, item, record, e) {
                record && me.applyBackground(record.get('spec'));
                if (e.type !== 'click') menu.hide();
            });
            menu.setInnerMenu([{ menu: this._bgPicker, index: 0 }]);
        },

        updateBackgrounds: function () {
            var me = this;
            this._bgStore && this._bgStore.each(function (m) { m.set('css', me.cssOf(m.get('spec'))); });
        },

        applyBackground: function (spec) {
            var api = this.api;
            if (!api) return;
            var pal = this.palette(), rgb = function (i) { return Common.Utils.ThemeColor.getRgbColor(pal[i] || { color: 'ffffff' }); };
            var props = new Asc.CAscSlideProps(), fill = new Asc.asc_CShapeFill();
            if (spec.length === 1) {
                fill.put_type(Asc.c_oAscFill.FILL_TYPE_SOLID);
                fill.put_fill(new Asc.asc_CFillSolid());
                fill.get_fill().put_color(rgb(spec[0]));
            } else {
                var g = new Asc.asc_CFillGrad();
                fill.put_type(Asc.c_oAscFill.FILL_TYPE_GRAD);
                fill.put_fill(g);
                g.put_grad_type(Asc.c_oAscFillGradType.GRAD_LINEAR);
                g.put_linear_angle(90 * 60000);
                g.put_linear_scale(true);
                g.put_positions([0, 100000]);
                g.put_colors([rgb(spec[0]), rgb(spec[1])]);
            }
            props.put_background(fill);
            api.SetSlideProps(props);
            var all = new Asc.CAscSlideProps();
            all.put_ApplyBackgroundToAll(true);
            api.SetSlideProps(all);
            Common.NotificationCenter.trigger('edit:complete', this.tb);
        }
    });
});

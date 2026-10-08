/*
 * SPDX-FileCopyrightText: 2026 ElixirOffice contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 *
 * ElixirOffice Design tab (document editor), like Word's:
 *   Document Formatting: Themes | Style Set gallery | Colors (existing) | Fonts | Paragraph Spacing
 *   Page Background: Watermark, Page Color (existing) | Page Borders
 * Themes apply an engine color scheme and the theme fonts (sdkjs asc_ElixirSetThemeFonts);
 * style sets and paragraph spacing rewrite the Title, Heading 1, Heading 2 and Normal styles
 * through the builder API (api.callCommand, one undo step); page borders use
 * asc_ElixirSetPageBorders. Presets: Presets.js. The only hooks are the controller and module
 * declarations in app.js / app_dev.js and an @import in app.less.
 */
define([
    'core',
    'common/main/lib/component/Window',
    'common/main/lib/component/ComboDataView',
    'common/main/lib/component/ComboBorderSize',
    'common/main/lib/component/ColorButton',
    'documenteditor/main/app/elixir-design/Presets'
], function () {
    'use strict';

    var PR = window.ElixirDesignPresets;
    var esc = function (s) { return Common.Utils.String.htmlEncode(s || ''); };

    DE.Controllers.ElixirDesign = Backbone.Controller.extend({

        initialize: function () {
            var me = this;
            this.schemes = [];
            Common.NotificationCenter.on('app:ready', function (mode) { _.defer(function () { me.onAppReady(mode); }); });
        },

        onLaunch: function () {},

        onAppReady: function (mode) {
            var main = DE.getController('Main'), tbc = DE.getController('Toolbar');
            if (!mode || !mode.isEdit || !main || !main.api || !tbc || !tbc.toolbar || this.ready) return;
            this.ready = true;
            this.api = main.api;
            this.toolbar = tbc.toolbar;
            var $panel = this.toolbar.$el.find('.panel[data-tab=design]');
            if (!$panel.length) return;
            this.render($panel);
            var me = this;
            this.api.asc_registerCallback('asc_onSendThemeColorSchemes', function (schemes) { me.schemes = schemes || []; me.fillThemes(); });
            this.fillThemes();
        },

        lock: function () {
            var _set = Common.enumLock;
            return [_set.previewReviewMode, _set.viewFormMode, _set.lostConnect, _set.disableOnStart, _set.docLockView, _set.docLockForms, _set.docLockComments, _set.viewMode];
        },

        render: function ($panel) {
            var me = this, tb = this.toolbar, $fmt = $panel.find('.group[data-group=docformatting]'), $bg = $panel.find('.group[data-group=pagebackground]');
            var slot = function (id) { return '<span class="btn-slot text x-huge" id="' + id + '"></span>'; };
            $fmt.prepend(slot('elx-dsg-themes'));
            $fmt.after('<div class="group small flex elx-dsg-stylesets" id="elx-dsg-stylesets" data-group="docformatting" style="width: 380px; min-width: 120px;" data-group-width="380px"></div>');
            var $tail = $('<div class="group" data-group="docformatting">' + slot('elx-dsg-fonts') + slot('elx-dsg-spacing') + '</div>');
            // Colors (the existing color scheme button) moves next to Fonts, after the gallery
            $panel.find('#elx-dsg-stylesets').after($tail);
            $tail.prepend($fmt.find('#slot-btn-colorschemas'));
            $bg.append(slot('elx-dsg-borders'));
            var btn = function (id, cfg) {
                var b = new Common.UI.Button(_.extend({cls: 'btn-toolbar x-huge icon-top', lock: me.lock(), dataHint: '1', dataHintDirection: 'bottom', dataHintOffset: 'small'}, cfg));
                b.render($panel.find('#' + id));
                tb.lockControls.push(b);
                return b;
            };
            this.btnThemes = btn('elx-dsg-themes', {iconCls: 'toolbar__icon btn-elx-themes', caption: this.txtThemes,
                menu: new Common.UI.Menu({cls: 'elx-dsg-theme-menu', style: 'width: 404px;', items: [
                    {template: _.template('<div id="elx-dsg-theme-list" class="elx-dsg-theme-list"></div>')}
                ]})});
            this.btnFonts = btn('elx-dsg-fonts', {iconCls: 'toolbar__icon btn-elx-theme-fonts', caption: this.txtFonts,
                menu: new Common.UI.Menu({cls: 'elx-dsg-fonts-menu', style: 'min-width: 240px;', items: []})});
            this.btnSpacing = btn('elx-dsg-spacing', {iconCls: 'toolbar__icon btn-elx-paragraph-spacing', caption: this.txtParagraphSpacing,
                menu: new Common.UI.Menu({cls: 'elx-dsg-spacing-menu', items: _.map(PR.spacing, function (s) {
                    return {caption: me['txtSpacing_' + s[0]], value: s[0], checkable: true, toggleGroup: 'elx-dsg-spacing'};
                })})});
            this.btnBorders = btn('elx-dsg-borders', {iconCls: 'toolbar__icon btn-elx-page-borders', caption: this.txtPageBorders});

            this.styleSets = new Common.UI.ComboDataView({
                cls: 'combo-styles elx-combo-stylesets',
                itemWidth: 86, itemHeight: 58,
                style: 'min-width: 120px; max-width: 370px;',
                lock: this.lock(),
                enableKeyEvents: true,
                autoWidth: true,
                menuMaxHeight: 300,
                delayRenderTips: true,
                ariaLabel: this.txtStyleSets,
                itemTemplate: _.template('<div class="elx-styleset elx-styleset-<%= key %>" id="<%= id %>">' +
                    '<div class="t">' + this.txtSampleTitle + '</div><div class="h">' + this.txtSampleHeading + '</div>' +
                    '<div class="b"></div><div class="b"></div><div class="b short"></div></div>'),
                dataHint: '1', dataHintDirection: 'bottom', dataHintOffset: '-16, 0'
            });
            this.styleSets.render($panel.find('#elx-dsg-stylesets'));
            this.styleSets.menuPicker.store.reset(_.map(PR.styleSets, function (s) { return {key: s.key, data: s.key, tip: me['txtSet_' + s.key] || s.name}; }));
            this.styleSets.fillComboView(this.styleSets.menuPicker.store.at(0), true);
            tb.lockControls.push(this.styleSets);
            if (tb.txtGroupStylesets === undefined) tb.txtGroupStylesets = this.txtStyleSets;

            this.styleSets.on('click', function (combo, record) { me.applyStyleSet(record.get('key')); });
            this.btnSpacing.menu.on('item:click', function (menu, item) { me.applySpacing(item.value); });
            this.btnFonts.menu.on('show:before', function (menu) { me.fillFonts(menu); });
            this.btnFonts.menu.on('item:click', function (menu, item) {
                if (item.value === 'custom') me.customFonts();
                else if (item.options.pair) me.applyFonts(item.options.pair[0], item.options.pair[1]);
            });
            this.btnThemes.menu.on('show:after', function () { me.fillThemes(); });
            $(document).on('click', '#elx-dsg-theme-list .elx-dsg-theme', function () {
                me.applyTheme(+$(this).attr('data-index'));
                me.btnThemes.menu.hide();
            });
            this.btnBorders.on('click', function () { me.pageBorders(); });
            Common.Utils.lockControls(Common.enumLock.disableOnStart, false, {array: [this.btnThemes, this.btnFonts, this.btnSpacing, this.btnBorders, this.styleSets]});
        },

        done: function () { Common.NotificationCenter.trigger('edit:complete', this.toolbar); },

        // ---- Themes -------------------------------------------------------------------------------
        schemeIndex: function (name) {
            return _.findIndex(this.schemes, function (s) { return s.get_name && s.get_name() === name; });
        },

        fillThemes: function () {
            var $list = $('#elx-dsg-theme-list');
            if (!this.schemes.length && this.api && this.api.getColorSchemes) {
                // the list was sent before this controller listened to it
                try { this.schemes = this.api.getColorSchemes(this.api.getCurrentTheme ? this.api.getCurrentTheme() : undefined).schemes || []; } catch (e) {}
            }
            if (!$list.length || !this.schemes.length) return;
            var me = this, cur = this.api.asc_ElixirGetThemeFonts && this.api.asc_ElixirGetThemeFonts(), html = [];
            _.each(PR.themes, function (t, i) {
                var k = me.schemeIndex(t[1]);
                if (k < 0) return;
                var colors = me.schemes[k].get_colors(), sw = '';
                for (var j = 4; j < 10; j++) {
                    var c = colors[j];
                    c && (sw += '<span style="background:#' + Common.Utils.ThemeColor.getHexColor(c.get_r(), c.get_g(), c.get_b()) + ';"></span>');
                }
                var on = cur && cur.Major === t[2] && cur.Minor === t[3];
                html.push('<div class="elx-dsg-theme' + (on ? ' selected' : '') + '" data-index="' + i + '" role="menuitem" title="' + esc(me.themeName(t[0])) + '">' +
                    '<div class="aa" style="font-family:\'' + esc(t[2]) + '\';">Aa</div><div class="sw">' + sw + '</div><div class="nm">' + esc(me.themeName(t[0])) + '</div></div>');
            });
            $list.html(html.join(''));
        },

        themeName: function (n) { return n === 'Office' ? this.txtOffice : n; },

        applyTheme: function (i) {
            var t = PR.themes[i], k = t && this.schemeIndex(t[1]);
            if (!t || k < 0) return;
            this.api.asc_ChangeColorSchemeByIdx(k);
            this.api.asc_ElixirSetThemeFonts(t[2], t[3]);
            this.done();
        },

        // ---- Fonts --------------------------------------------------------------------------------
        fillFonts: function (menu) {
            var me = this, cur = this.api.asc_ElixirGetThemeFonts && this.api.asc_ElixirGetThemeFonts();
            menu.removeAll();
            var tpl = _.template('<a id="<%= id %>" tabindex="-1" type="menuitem" class="elx-dsg-fontpair"><span class="nm"><%- caption %></span>' +
                '<span class="mj" style="font-family:\'<%- options.pair[0] %>\';"><%- options.pair[0] %></span>' +
                '<span class="mn" style="font-family:\'<%- options.pair[1] %>\';"><%- options.pair[1] %></span></a>');
            menu.addItem({caption: this.txtOffice, template: _.template('<div class="elx-dsg-menu-header">' + esc(this.txtOffice) + '</div>'), disabled: true});
            _.each(PR.fonts, function (f) {
                menu.addItem({template: tpl, caption: f[0], pair: [f[1], f[2]], checkable: true, toggleGroup: 'elx-dsg-fonts',
                    checked: !!cur && cur.Major === f[1] && cur.Minor === f[2]});
            });
            menu.addItem({caption: '--'});
            menu.addItem({caption: this.txtCustomizeFonts, value: 'custom'});
        },

        applyFonts: function (major, minor) {
            this.api.asc_ElixirSetThemeFonts(major, minor);
            this.done();
        },

        fontNames: function () {
            var tb = this.toolbar, store = tb.cmbFontName && tb.cmbFontName.store;
            var names = store ? store.map(function (m) { return m.get('name'); }) : [];
            return _.uniq(names).sort();
        },

        customFonts: function () {
            var me = this, cur = this.api.asc_ElixirGetThemeFonts() || {Major: 'Calibri Light', Minor: 'Calibri'},
                names = this.fontNames();
            var win = new Common.UI.Window({
                title: this.txtCreateFonts, width: 380, cls: 'modal-dlg elx-dsg-fonts-dlg',
                tpl: '<div class="box" style="padding: 8px 4px;">' +
                    '<label class="input-label">' + esc(this.txtHeadingFont) + '</label><div id="elx-dsg-font-major" style="margin-bottom: 12px;"></div>' +
                    '<label class="input-label">' + esc(this.txtBodyFont) + '</label><div id="elx-dsg-font-minor"></div>' +
                    '</div><div class="footer center"><button class="btn normal dlg-btn primary" result="ok">' + esc(this.txtSave) + '</button>' +
                    '<button class="btn normal dlg-btn" result="cancel">' + esc(this.txtCancel) + '</button></div>'
            });
            // the controls are built when the window renders, so that it is sized and centered with them
            win.on('render:after', function () {
                var combo = function (el, value) {
                    var c = new Common.UI.ComboBox({el: win.$window.find(el), cls: 'input-group-nr', style: 'width: 100%;', menuStyle: 'min-width: 100%; max-height: 220px;', editable: true,
                        data: _.map(names, function (n) { return {value: n, displayValue: n}; })});
                    c.setValue(value);
                    return c;
                };
                var cMajor = combo('#elx-dsg-font-major', cur.Major), cMinor = combo('#elx-dsg-font-minor', cur.Minor);
                var val = function (c) { var v = c.getRawValue ? c.getRawValue() : c.getValue(); return (v || '').trim(); };
                win.on('close', function () { me.done(); });
                win.$window.find('.dlg-btn').on('click', function () {
                    if ($(this).attr('result') === 'ok' && val(cMajor) && val(cMinor)) me.api.asc_ElixirSetThemeFonts(val(cMajor), val(cMinor));
                    win.close();
                });
            });
            win.show();
        },

        // ---- Style sets and paragraph spacing (builder API, one undo step) ---------------------------
        accentHex: function () {
            var k = -1, eff = Common.Utils.ThemeColor.getEffectColors && Common.Utils.ThemeColor.getEffectColors();
            if (eff && eff.length > 4) return (eff[4].color || eff[4]).toString().replace('#', '').toUpperCase();
            return '4472C4';
        },

        run: function (source) {
            var me = this;
            if (!this.api || !this.api.callCommand) return;
            this.api.callCommand(source, true, function () { setTimeout(function () { me.done(); }, 0); });
        },

        applyStyleSet: function (key) {
            var set = _.findWhere(PR.styleSets, {key: key});
            if (!set) return;
            var spec = _.omit(set, 'key', 'name');
            this.run('(function () {' +
                'var doc = Api.GetDocument(), spec = ' + JSON.stringify(spec) + ', accent = "' + this.accentHex() + '";' +
                'var rgb = function (h) { return [parseInt(h.substr(0, 2), 16), parseInt(h.substr(2, 2), 16), parseInt(h.substr(4, 2), 16)]; };' +
                'Object.keys(spec).forEach(function (name) {' +
                '  var st = doc.GetStyle(name); if (!st) return; var s = spec[name], tp = st.GetTextPr(), pp = st.GetParaPr();' +
                '  if (s.size) tp.SetFontSize(s.size * 2);' +
                '  tp.SetBold(!!s.bold); tp.SetItalic(!!s.italic); tp.SetCaps(!!s.caps);' +
                '  if (s.color === "accent") tp.SetColor(Api.ThemeColor("accent1"));' +
                '  else if (s.color) { var c = rgb(s.color); tp.SetColor(c[0], c[1], c[2]); }' +
                '  else tp.SetColor(Api.ThemeColor("tx1"));' +
                '  pp.SetJc(s.align === "center" ? "center" : "left");' +
                '  if (s.before !== undefined) pp.SetSpacingBefore(s.before * 20);' +
                '  if (s.after !== undefined) pp.SetSpacingAfter(s.after * 20);' +
                '  if (s.border) { var b = rgb(s.border === "accent" ? accent : s.border); pp.SetBottomBorder("single", 6, 1, b[0], b[1], b[2]); }' +
                '  else pp.SetBottomBorder("none", 0, 0, 0, 0, 0);' +
                '  if (s.shading) { if (s.shading === "accent") pp.SetShd("clear", Api.ThemeColor("accent1")); else { var d = rgb(s.shading); pp.SetShd("clear", d[0], d[1], d[2]); } }' +
                '  else pp.SetShd("nil", 255, 255, 255, true);' +
                '});' +
                'return true; })()');
            this._styleSet = key;
        },

        applySpacing: function (key) {
            var s = _.find(PR.spacing, function (x) { return x[0] === key; });
            if (!s) return;
            this.run('(function () {' +
                'var doc = Api.GetDocument(), st = doc.GetStyle("Normal"), pp = st ? st.GetParaPr() : doc.GetDefaultParaPr();' +
                'pp.SetSpacingBefore(' + (s[1] * 20) + '); pp.SetSpacingAfter(' + (s[2] * 20) + '); pp.SetSpacingLine(' + Math.round(s[3] * 240) + ', "auto");' +
                'return true; })()');
        },

        // ---- Page Borders -------------------------------------------------------------------------
        pageBorders: function () {
            var me = this, cur = this.api.asc_ElixirGetPageBorders() || {Value: 0};
            var win = new Common.UI.Window({
                title: this.txtPageBordersTitle, width: 360, cls: 'modal-dlg elx-dsg-borders-dlg',
                tpl: '<div class="box" style="padding: 8px 4px;">' +
                    '<label class="input-label">' + esc(this.txtSetting) + '</label>' +
                    '<div class="elx-dsg-setting" style="display:flex; gap: 8px; margin-bottom: 12px;">' +
                    '<button class="btn normal elx-dsg-none" data-v="0">' + esc(this.txtNone) + '</button>' +
                    '<button class="btn normal elx-dsg-box" data-v="1">' + esc(this.txtBox) + '</button></div>' +
                    '<div style="display:flex; gap: 16px; margin-bottom: 12px;">' +
                    '<div><label class="input-label">' + esc(this.txtColor) + '</label><div id="elx-dsg-bcolor"></div></div>' +
                    '<div><label class="input-label">' + esc(this.txtWidth) + '</label><div id="elx-dsg-bsize"></div></div></div>' +
                    '<label class="input-label">' + esc(this.txtMeasureFrom) + '</label><div id="elx-dsg-bfrom"></div>' +
                    '</div><div class="footer center"><button class="btn normal dlg-btn primary" result="ok">' + esc(this.txtOk) + '</button>' +
                    '<button class="btn normal dlg-btn" result="cancel">' + esc(this.txtCancel) + '</button></div>'
            });
            // the controls are built when the window renders, so that it is sized and centered with them
            win.on('render:after', function () {
                var state = {value: cur.Value ? 1 : 0, color: cur.Color ? Common.Utils.ThemeColor.getHexColor(cur.Color.r, cur.Color.g, cur.Color.b) : '000000'};
                var $set = win.$window.find('.elx-dsg-setting .btn');
                var sync = function () { $set.each(function () { $(this).toggleClass('active', +$(this).attr('data-v') === state.value); }); };
                $set.on('click', function () { state.value = +$(this).attr('data-v'); sync(); });
                sync();
                var color = new Common.UI.ColorButton({parentEl: win.$window.find('#elx-dsg-bcolor'), color: state.color, menu: true, eyeDropper: false});
                color.on('color:select', function (btn, c) { state.color = typeof c === 'object' ? c.color : c; state.value = 1; sync(); });
                var size = new Common.UI.ComboBorderSize({el: win.$window.find('#elx-dsg-bsize'), style: 'width: 110px;'});
                size.setValue(cur.Value && cur.Size ? cur.Size * 72 / 25.4 : 0.5);
                if (size.getValue() === undefined || size.getValue() === null || size.getValue() === '') size.selectRecord(size.store.at(1));
                var from = new Common.UI.ComboBox({el: win.$window.find('#elx-dsg-bfrom'), cls: 'input-group-nr', style: 'width: 200px;', editable: false,
                    data: [{value: 0, displayValue: me.txtFromText}, {value: 1, displayValue: me.txtFromEdge}]});
                from.setValue(cur.OffsetFrom === 1 ? 1 : 0);
                win.on('close', function () { me.done(); });
                win.$window.find('.dlg-btn').on('click', function () {
                    if ($(this).attr('result') === 'ok') {
                        var hex = (state.color || '000000').replace('#', ''), pt = parseFloat(size.getValue()) || 0.5, edge = from.getValue() === 1;
                        me.api.asc_ElixirSetPageBorders({
                            Value: state.value, Size: pt * 25.4 / 72, Space: (edge ? 24 : 4) * 25.4 / 72, OffsetFrom: edge ? 1 : 0,
                            Color: {r: parseInt(hex.substr(0, 2), 16), g: parseInt(hex.substr(2, 2), 16), b: parseInt(hex.substr(4, 2), 16)}
                        });
                    }
                    win.close();
                });
            });
            win.show();
        },

        // ---- strings (English; translations in locale/*.json) ------------------------------------
        txtThemes: 'Themes',
        txtOffice: 'Office',
        txtStyleSets: 'Style Sets',
        txtSampleTitle: 'Title',
        txtSampleHeading: 'Heading 1',
        txtFonts: 'Fonts',
        txtCustomizeFonts: 'Customize Fonts...',
        txtCreateFonts: 'Create New Theme Fonts',
        txtHeadingFont: 'Heading font:',
        txtBodyFont: 'Body font:',
        txtSave: 'Save',
        txtCancel: 'Cancel',
        txtOk: 'OK',
        txtParagraphSpacing: 'Paragraph Spacing',
        txtSpacing_default: 'Default',
        txtSpacing_none: 'No Paragraph Space',
        txtSpacing_compact: 'Compact',
        txtSpacing_tight: 'Tight',
        txtSpacing_open: 'Open',
        txtSpacing_relaxed: 'Relaxed',
        txtSpacing_double: 'Double',
        txtSet_basic: 'Basic (Simple)',
        txtSet_elegant: 'Basic (Elegant)',
        txtSet_stylish: 'Basic (Stylish)',
        txtSet_casual: 'Casual',
        txtSet_centered: 'Centered',
        txtSet_lines: 'Lines (Simple)',
        txtSet_distinctive: 'Lines (Distinctive)',
        txtSet_shaded: 'Shaded',
        txtSet_minimalist: 'Minimalist',
        txtPageBorders: 'Page Borders',
        txtPageBordersTitle: 'Page Borders',
        txtSetting: 'Setting:',
        txtNone: 'None',
        txtBox: 'Box',
        txtColor: 'Color:',
        txtWidth: 'Width:',
        txtMeasureFrom: 'Measure from:',
        txtFromText: 'Text',
        txtFromEdge: 'Edge of page'
    });
});

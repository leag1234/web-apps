/*
 * SPDX-FileCopyrightText: 2026 ElixirOffice contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 *
 * DE.Views.FontDialog: Word's Font dialog (Font group launcher, Ctrl+D).
 *  - Font tab: font, font style and size (text box over a list, as in Word), font colour,
 *    underline style and colour, effects (strikethrough, double strikethrough, superscript,
 *    subscript, small caps, all caps, hidden);
 *  - Advanced tab: character spacing (scale, spacing, position, kerning) and OpenType ligatures.
 * Values come from and go to api.asc_GetFontDialogProps / asc_SetFontDialogProps (one undo step);
 * only the values the user changed are applied, so a mixed selection keeps its other formatting.
 * In styleMode (Modify Style > Format > Font...) the values are returned to the handler instead.
 * The engine has no underline styles other than single, double and words only, no underline colour, character scale or
 * kerning: those controls show the Word layout but are disabled.
 */

define([
    'common/main/lib/component/Window',
    'common/main/lib/component/ComboBox',
    'common/main/lib/component/ListView',
    'common/main/lib/component/MetricSpinner',
    'common/main/lib/component/CheckBox'
], function () { 'use strict';

    var PT_TO_MM = 25.4 / 72;
    var SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72];

    DE.Views.FontDialog = Common.UI.Window.extend(_.extend({
        options: {
            width: 470,
            cls: 'modal-dlg word-font-dlg',
            id: 'window-word-font',
            buttons: ['ok', 'cancel']
        },

        initialize: function (options) {
            _.extend(this.options, {title: this.txtTitle}, options || {});
            var t = this;
            this.options.tpl = [
                '<div class="box">',
                    '<div class="word-dlg-tabs" role="tablist">',
                        '<button type="button" class="word-dlg-tab active" data-tab="font" role="tab">' + t.txtFontTab + '</button>',
                        '<button type="button" class="word-dlg-tab" data-tab="advanced" role="tab">' + t.txtAdvancedTab + '</button>',
                    '</div>',
                    '<div class="word-dlg-panel" data-panel="font">',
                        '<table class="word-font-lists"><tbody><tr>',
                            '<td class="col-font"><label>' + t.txtFont + '</label><div class="wf-font-input"></div><div class="wf-font-list"></div></td>',
                            '<td class="col-style"><label>' + t.txtFontStyle + '</label><div class="wf-style-input"></div><div class="wf-style-list"></div></td>',
                            '<td class="col-size"><label>' + t.txtSize + '</label><div class="wf-size-input"></div><div class="wf-size-list"></div></td>',
                        '</tr></tbody></table>',
                        '<table class="word-font-row2"><tbody><tr>',
                            '<td><label>' + t.txtFontColor + '</label><div class="wf-color"></div></td>',
                            '<td><label>' + t.txtUnderlineStyle + '</label><div class="wf-underline input-group-nr"></div></td>',
                            '<td><label>' + t.txtUnderlineColor + '</label><div class="wf-underline-color"></div></td>',
                        '</tr></tbody></table>',
                        '<fieldset class="word-dlg-group"><legend>' + t.txtEffects + '</legend>',
                            '<table class="word-font-effects"><tbody>',
                                '<tr><td><div class="wf-strike"></div></td><td><div class="wf-smallcaps"></div></td></tr>',
                                '<tr><td><div class="wf-dstrike"></div></td><td><div class="wf-allcaps"></div></td></tr>',
                                '<tr><td><div class="wf-super"></div></td><td><div class="wf-hidden"></div></td></tr>',
                                '<tr><td><div class="wf-sub"></div></td><td></td></tr>',
                            '</tbody></table>',
                        '</fieldset>',
                    '</div>',
                    '<div class="word-dlg-panel" data-panel="advanced" style="display: none;">',
                        '<fieldset class="word-dlg-group"><legend>' + t.txtCharSpacing + '</legend>',
                            '<table class="word-font-adv"><tbody>',
                                '<tr><td><label>' + t.txtScale + '</label></td><td><div class="wf-scale input-group-nr"></div></td><td></td><td></td></tr>',
                                '<tr><td><label>' + t.txtSpacing + '</label></td><td><div class="wf-spacing-type input-group-nr"></div></td><td><label>' + t.txtBy + '</label></td><td><div class="wf-spacing-by"></div></td></tr>',
                                '<tr><td><label>' + t.txtPosition + '</label></td><td><div class="wf-position-type input-group-nr"></div></td><td><label>' + t.txtBy + '</label></td><td><div class="wf-position-by"></div></td></tr>',
                                '<tr><td colspan="2"><div class="wf-kerning"></div></td><td colspan="2"><div class="wf-kerning-by"></div><label class="wf-kerning-label">' + t.txtPointsAbove + '</label></td></tr>',
                            '</tbody></table>',
                        '</fieldset>',
                        '<fieldset class="word-dlg-group"><legend>' + t.txtOpenType + '</legend>',
                            '<table class="word-font-adv"><tbody>',
                                '<tr><td><label>' + t.txtLigatures + '</label></td><td><div class="wf-ligatures input-group-nr"></div></td></tr>',
                                '<tr><td><label>' + t.txtNumSpacing + '</label></td><td><div class="wf-numspacing input-group-nr"></div></td></tr>',
                                '<tr><td><label>' + t.txtNumForms + '</label></td><td><div class="wf-numforms input-group-nr"></div></td></tr>',
                                '<tr><td><label>' + t.txtStylisticSets + '</label></td><td><div class="wf-stylistic input-group-nr"></div></td></tr>',
                            '</tbody></table>',
                        '</fieldset>',
                    '</div>',
                    '<div class="word-font-preview-title">' + t.txtPreview + '</div>',
                    '<div class="word-font-preview" id="word-font-dlg-preview"></div>',
                    '<div class="word-font-note"></div>',
                '</div>'
            ].join('');
            this.api = options.api;
            this.fontStore = options.fontStore;
            this.handler = options.handler;
            this.styleMode = !!options.styleMode;
            this.props = _.clone(options.props || {});
            this.previewText = options.previewText;
            this.changes = {};
            Common.UI.Window.prototype.initialize.call(this, this.options);
        },

        render: function () {
            Common.UI.Window.prototype.render.call(this);
            var me = this, $w = this.getChild(), t = this, p = this.props;

            $w.find('.dlg-btn').on('click', _.bind(this.onBtnClick, this));
            $w.find('.word-dlg-tab').on('click', function (e) { me.selectTab($(e.currentTarget).attr('data-tab')); });

            // ---- font / style / size: text box over a list
            this.inputFont = new Common.UI.InputField({el: $w.find('.wf-font-input'), style: 'width: 100%;', value: p.FontName || ''});
            this.listFont = new Common.UI.ListView({
                el: $w.find('.wf-font-list'), store: new Common.UI.DataViewStore(), tabindex: 1, cls: 'word-font-listview',
                itemTemplate: _.template('<div id="<%= id %>" class="list-item"><%= Common.Utils.String.htmlEncode(value) %></div>')
            });
            var fonts = [];
            this.fontStore && this.fontStore.each(function (rec) {
                var name = rec.get('name');
                name && !rec.get('cloneid') && fonts.push({id: Common.UI.getId(), value: name});
            });
            fonts.sort(function (a, b) { return a.value.localeCompare(b.value, undefined, {sensitivity: 'base'}); });
            this.listFont.store.reset(fonts);
            this.listFont.on('item:select', function (view, item, rec) {
                me.inputFont.setValue(rec.get('value'));
                me.change('FontName', rec.get('value'));
            });
            this.inputFont._input.on('input', function () { me.findInList(me.listFont, me.inputFont.getValue(), function (v) { me.change('FontName', v); }); });

            this.styles = [
                {id: Common.UI.getId(), value: t.txtRegular, bold: false, italic: false},
                {id: Common.UI.getId(), value: t.txtItalic, bold: false, italic: true},
                {id: Common.UI.getId(), value: t.txtBold, bold: true, italic: false},
                {id: Common.UI.getId(), value: t.txtBoldItalic, bold: true, italic: true}
            ];
            this.inputStyle = new Common.UI.InputField({el: $w.find('.wf-style-input'), style: 'width: 100%;'});
            this.listStyle = new Common.UI.ListView({
                el: $w.find('.wf-style-list'), store: new Common.UI.DataViewStore(this.styles), tabindex: 1, cls: 'word-font-listview',
                itemTemplate: _.template('<div id="<%= id %>" class="list-item"><%= Common.Utils.String.htmlEncode(value) %></div>')
            });
            this.listStyle.on('item:select', function (view, item, rec) {
                me.inputStyle.setValue(rec.get('value'));
                me.change('Bold', rec.get('bold'));
                me.change('Italic', rec.get('italic'));
            });
            this.inputStyle._input.on('input', function () {
                me.findInList(me.listStyle, me.inputStyle.getValue(), function (v, rec) { me.change('Bold', rec.get('bold')); me.change('Italic', rec.get('italic')); });
            });

            this.inputSize = new Common.UI.InputField({el: $w.find('.wf-size-input'), style: 'width: 100%;', value: p.FontSize ? String(p.FontSize) : ''});
            this.listSize = new Common.UI.ListView({
                el: $w.find('.wf-size-list'), store: new Common.UI.DataViewStore(_.map(SIZES, function (s) { return {id: Common.UI.getId(), value: String(s), size: s}; })),
                tabindex: 1, cls: 'word-font-listview',
                itemTemplate: _.template('<div id="<%= id %>" class="list-item"><%= value %></div>')
            });
            this.listSize.on('item:select', function (view, item, rec) {
                me.inputSize.setValue(rec.get('value'));
                me.change('FontSize', rec.get('size'));
            });
            this.inputSize._input.on('input', function () {
                var v = parseFloat(String(me.inputSize.getValue()).replace(',', '.'));
                if (v >= 1 && v <= 1638) {
                    me.change('FontSize', Math.round(v * 2) / 2);
                    var rec = me.listSize.store.findWhere({size: v});
                    rec ? me.listSize.selectRecord(rec, true) : me.listSize.deselectAll(true);
                    rec && me.listSize.scrollToRecord(rec);
                }
            });

            // ---- colour, underline
            this.btnColor = new Common.UI.ButtonColored({
                parentEl: $w.find('.wf-color'), cls: 'btn-toolbar word-font-colorbtn', iconCls: 'toolbar__icon btn-fontcolor',
                hint: t.txtFontColor, auto: true, color: '000000', menu: true
            });
            this.btnColor.setMenu();
            this.mnuColor = this.btnColor.getPicker();
            this.btnColor.on('color:select', function (b, color) { me.change('Color', Common.Utils.ThemeColor.getRgbColor(color)); });
            this.btnColor.on('auto:select', function () { var c = new Asc.asc_CColor(); c.put_auto(true); me.change('Color', c); });

            this.cmbUnderline = new Common.UI.ComboBox({
                el: $w.find('.wf-underline'), cls: 'input-group-nr', editable: false, menuStyle: 'min-width: 100%;',
                data: [{value: 0, displayValue: t.txtNone}, {value: 1, displayValue: t.txtUnderlineSingle},
                    {value: 10, displayValue: t.txtUnderlineDouble}, {value: 17, displayValue: t.txtUnderlineWords}]
            });
            this.cmbUnderline.on('selected', function (c, r) { me.change('Underline', r.value > 1 ? r.value : r.value === 1); });
            this.btnUnderlineColor = new Common.UI.Button({
                parentEl: $w.find('.wf-underline-color'), cls: 'btn-text-menu-default word-font-ucolor', caption: t.txtAutomatic, disabled: true,
                hint: t.txtNotSupported
            });

            // ---- effects
            var chk = function (cls, label, key) {
                var c = new Common.UI.CheckBox({el: $w.find(cls), labelText: label});
                c.on('change', function (field, value) { me.onEffect(key, value === 'checked' || value === true); });
                return c;
            };
            this.chStrike = chk('.wf-strike', t.txtStrike, 'Strikeout');
            this.chDStrike = chk('.wf-dstrike', t.txtDStrike, 'DStrikeout');
            this.chSuper = chk('.wf-super', t.txtSuperscript, 'Superscript');
            this.chSub = chk('.wf-sub', t.txtSubscript, 'Subscript');
            this.chSmallCaps = chk('.wf-smallcaps', t.txtSmallCaps, 'SmallCaps');
            this.chAllCaps = chk('.wf-allcaps', t.txtAllCaps, 'Caps');
            this.chHidden = chk('.wf-hidden', t.txtHidden, 'Vanish');

            // ---- advanced
            this.cmbScale = new Common.UI.ComboBox({
                el: $w.find('.wf-scale'), cls: 'input-group-nr', editable: false, disabled: true, menuStyle: 'min-width: 100%;',
                data: [{value: 100, displayValue: '100%'}]
            });
            this.cmbScale.setValue(100);
            this.cmbSpacing = new Common.UI.ComboBox({
                el: $w.find('.wf-spacing-type'), cls: 'input-group-nr', editable: false, menuStyle: 'min-width: 100%;',
                data: [{value: 0, displayValue: t.txtNormal}, {value: 1, displayValue: t.txtExpanded}, {value: -1, displayValue: t.txtCondensed}]
            });
            this.numSpacing = new Common.UI.MetricSpinner({
                el: $w.find('.wf-spacing-by'), step: .1, width: 80, defaultUnit: 'pt', value: '0 pt', maxValue: 1584, minValue: 0
            });
            this.cmbPosition = new Common.UI.ComboBox({
                el: $w.find('.wf-position-type'), cls: 'input-group-nr', editable: false, menuStyle: 'min-width: 100%;',
                data: [{value: 0, displayValue: t.txtNormal}, {value: 1, displayValue: t.txtRaised}, {value: -1, displayValue: t.txtLowered}]
            });
            this.numPosition = new Common.UI.MetricSpinner({
                el: $w.find('.wf-position-by'), step: .5, width: 80, defaultUnit: 'pt', value: '0 pt', maxValue: 1584, minValue: 0
            });
            var spacingChanged = function () {
                var type = me.cmbSpacing.getValue(), by = me.numSpacing.getNumberValue() || 0;
                if (type === 0) by = 0;
                else if (by === 0) { by = 1; me.numSpacing.setValue(1, true); }
                me.change('Spacing', type * by * PT_TO_MM);
            };
            this.cmbSpacing.on('selected', spacingChanged);
            this.numSpacing.on('change', function () {
                if (me.cmbSpacing.getValue() === 0 && me.numSpacing.getNumberValue() > 0) me.cmbSpacing.setValue(1);
                spacingChanged();
            });
            var positionChanged = function () {
                var type = me.cmbPosition.getValue(), by = me.numPosition.getNumberValue() || 0;
                if (type === 0) by = 0;
                else if (by === 0) { by = 3; me.numPosition.setValue(3, true); }
                me.change('Position', type * by * PT_TO_MM);
            };
            this.cmbPosition.on('selected', positionChanged);
            this.numPosition.on('change', function () {
                if (me.cmbPosition.getValue() === 0 && me.numPosition.getNumberValue() > 0) me.cmbPosition.setValue(1);
                positionChanged();
            });
            this.chKerning = new Common.UI.CheckBox({el: $w.find('.wf-kerning'), labelText: t.txtKerning, disabled: true});
            this.numKerning = new Common.UI.MetricSpinner({
                el: $w.find('.wf-kerning-by'), step: .5, width: 60, defaultUnit: 'pt', value: '', maxValue: 1638, minValue: 1, disabled: true
            });
            var L = Asc.LigaturesType || {};
            this.cmbLigatures = new Common.UI.ComboBox({
                el: $w.find('.wf-ligatures'), cls: 'input-group-nr', editable: false, menuStyle: 'min-width: 100%;',
                data: [{value: L.None, displayValue: t.txtNone}, {value: L.Standard, displayValue: t.txtLigStandard},
                       {value: L.StandardContextual, displayValue: t.txtLigStandardContextual},
                       {value: L.HistoricalDiscretional, displayValue: t.txtLigHistorical}, {value: L.All, displayValue: t.txtLigAll}]
            });
            this.cmbLigatures.on('selected', function (c, r) { me.change('Ligatures', r.value); });
            _.each(['.wf-numspacing', '.wf-numforms', '.wf-stylistic'], function (cls) {
                var c = new Common.UI.ComboBox({el: $w.find(cls), cls: 'input-group-nr', editable: false, disabled: true, menuStyle: 'min-width: 100%;',
                    data: [{value: 0, displayValue: t.txtDefault}]});
                c.setValue(0);
            });

            this.setValues();
        },

        show: function () {
            Common.UI.Window.prototype.show.apply(this, arguments);
            var me = this;
            this.mnuColor.updateColors(Common.Utils.ThemeColor.getEffectColors(), Common.Utils.ThemeColor.getStandartColors());
            this.setColorButton(this.props.Color);
            setTimeout(function () {
                var rec = me.props.FontName ? me.listFont.store.findWhere({value: me.props.FontName}) : null;
                rec && (me.listFont.selectRecord(rec, true), me.listFont.scrollToRecord(rec, true));
                var srec = me.listSize.getSelectedRec();
                srec && me.listSize.scrollToRecord(srec, true);
                me.updatePreview();
                me.inputFont.focus();
                me.inputFont._input.select();
            }, 150);
        },

        selectTab: function (tab) {
            var $w = this.getChild();
            $w.find('.word-dlg-tab').removeClass('active').filter('[data-tab="' + tab + '"]').addClass('active');
            $w.find('.word-dlg-panel').hide().filter('[data-panel="' + tab + '"]').show();
            this.updatePreview();
        },

        // select the first list item starting with the typed text (Word's font box behaviour)
        findInList: function (list, text, onFound) {
            text = (text || '').trim().toLowerCase();
            if (!text) return;
            var rec = list.store.find(function (r) { return r.get('value').toLowerCase() === text; }) ||
                      list.store.find(function (r) { return r.get('value').toLowerCase().indexOf(text) === 0; });
            if (rec) {
                list.selectRecord(rec, true);
                list.scrollToRecord(rec);
                rec.get('value').toLowerCase() === text && onFound(rec.get('value'), rec);
            }
        },

        setValues: function () {
            var p = this.props, me = this;
            var st = _.find(this.styles, function (s) { return s.bold === !!p.Bold && s.italic === !!p.Italic; });
            if (p.Bold !== undefined && p.Italic !== undefined && st) {
                this.inputStyle.setValue(st.value);
                this.listStyle.selectRecord(this.listStyle.store.findWhere({value: st.value}), true);
            }
            var srec = p.FontSize ? this.listSize.store.findWhere({size: p.FontSize}) : null;
            srec && this.listSize.selectRecord(srec, true);
            this.cmbUnderline.setValue(p.Underline > 1 ? p.Underline : (p.Underline ? 1 : 0));
            var set = function (c, v) { c.setValue(v === undefined ? 'indeterminate' : !!v, true); };
            set(this.chStrike, p.Strikeout);
            set(this.chDStrike, p.DStrikeout);
            set(this.chSuper, p.VertAlign === undefined ? undefined : p.VertAlign === 2);
            set(this.chSub, p.VertAlign === undefined ? undefined : p.VertAlign === 1);
            set(this.chSmallCaps, p.SmallCaps);
            set(this.chAllCaps, p.Caps);
            set(this.chHidden, p.Vanish);
            var sp = (p.Spacing || 0) / PT_TO_MM;
            this.cmbSpacing.setValue(sp > 0.001 ? 1 : (sp < -0.001 ? -1 : 0));
            this.numSpacing.setValue(Math.abs(Math.round(sp * 10) / 10), true);
            var pos = (p.Position || 0) / PT_TO_MM;
            this.cmbPosition.setValue(pos > 0.001 ? 1 : (pos < -0.001 ? -1 : 0));
            this.numPosition.setValue(Math.abs(Math.round(pos * 10) / 10), true);
            var lig = this.cmbLigatures.store.findWhere({value: p.Ligatures});
            this.cmbLigatures.setValue(lig ? p.Ligatures : (Asc.LigaturesType ? Asc.LigaturesType.None : 0));
        },

        setColorButton: function (color) {
            if (!color || (color.get_auto && color.get_auto())) {
                this.btnColor.setAutoColor(true);
                this.mnuColor.clearSelection();
                return;
            }
            var clr = color.get_type() == Asc.c_oAscColor.COLOR_TYPE_SCHEME ?
                {color: Common.Utils.ThemeColor.getHexColor(color.get_r(), color.get_g(), color.get_b()), effectValue: color.get_value()} :
                Common.Utils.ThemeColor.getHexColor(color.get_r(), color.get_g(), color.get_b());
            this.btnColor.setAutoColor(false);
            this.btnColor.setColor(typeof clr === 'object' ? clr.color : clr);
            Common.Utils.ThemeColor.selectPickerColorByEffect(clr, this.mnuColor);
        },

        // Word: strikethrough/double, superscript/subscript and small caps/all caps exclude each other
        onEffect: function (key, value) {
            if (this._inEffect) return;
            this._inEffect = true;
            var pairs = {Strikeout: this.chDStrike, DStrikeout: this.chStrike, Superscript: this.chSub, Subscript: this.chSuper, SmallCaps: this.chAllCaps, Caps: this.chSmallCaps};
            if (value && pairs[key]) pairs[key].setValue(false, true);
            switch (key) {
                case 'Superscript':
                case 'Subscript':
                    var sup = this.chSuper.getValue() === 'checked', sub = this.chSub.getValue() === 'checked';
                    this.change('VertAlign', sup ? 2 : (sub ? 1 : 0)); // AscCommon.vertalign_SuperScript = 2, SubScript = 1
                    break;
                case 'Strikeout':
                case 'DStrikeout':
                    this.change('Strikeout', this.chStrike.getValue() === 'checked');
                    this.change('DStrikeout', this.chDStrike.getValue() === 'checked');
                    break;
                case 'SmallCaps':
                case 'Caps':
                    this.change('SmallCaps', this.chSmallCaps.getValue() === 'checked');
                    this.change('Caps', this.chAllCaps.getValue() === 'checked');
                    break;
                default:
                    this.change(key, value);
            }
            this._inEffect = false;
        },

        change: function (key, value) {
            this.changes[key] = value;
            this.props[key] = value;
            this.updatePreview();
        },

        updatePreview: function () {
            var me = this;
            clearTimeout(this._previewTimer);
            this._previewTimer = setTimeout(function () {
                if (!me.api || !me.api.asc_DrawFontDialogPreview || !me.getChild().is(':visible')) return;
                var text = me.previewText || me.props.FontName || 'AaBbYyZz';
                me.api.asc_DrawFontDialogPreview('word-font-dlg-preview', me.props, text);
            }, 30);
        },

        getChanges: function () { return _.clone(this.changes); },

        onBtnClick: function (event) {
            this._handleInput(event.currentTarget.attributes['result'].value);
        },

        onPrimary: function () {
            this._handleInput('ok');
            return false;
        },

        _handleInput: function (state) {
            this.handler && this.handler.call(this, state, state === 'ok' ? this.getChanges() : null);
            this.close();
        },

        getFocusedComponents: function () {
            return [this.inputFont, this.listFont, this.inputStyle, this.listStyle, this.inputSize, this.listSize, this.btnColor, this.cmbUnderline,
                this.chStrike, this.chDStrike, this.chSuper, this.chSub, this.chSmallCaps, this.chAllCaps, this.chHidden,
                this.cmbSpacing, this.numSpacing, this.cmbPosition, this.numPosition, this.cmbLigatures].concat(this.getFooterButtons());
        },

        getDefaultFocusableComponent: function () {
            return this.inputFont;
        },

        txtTitle: 'Font',
        txtFontTab: 'Font',
        txtAdvancedTab: 'Advanced',
        txtFont: 'Font:',
        txtFontStyle: 'Font style:',
        txtSize: 'Size:',
        txtRegular: 'Regular',
        txtItalic: 'Italic',
        txtBold: 'Bold',
        txtBoldItalic: 'Bold Italic',
        txtFontColor: 'Font color:',
        txtUnderlineStyle: 'Underline style:',
        txtUnderlineColor: 'Underline color:',
        txtNone: '(none)',
        txtUnderlineSingle: 'Single',
        txtUnderlineDouble: 'Double',
        txtUnderlineWords: 'Words only',
        txtAutomatic: 'Automatic',
        txtNotSupported: 'Not available in this version',
        txtEffects: 'Effects',
        txtStrike: 'Strikethrough',
        txtDStrike: 'Double strikethrough',
        txtSuperscript: 'Superscript',
        txtSubscript: 'Subscript',
        txtSmallCaps: 'Small caps',
        txtAllCaps: 'All caps',
        txtHidden: 'Hidden',
        txtCharSpacing: 'Character Spacing',
        txtScale: 'Scale:',
        txtSpacing: 'Spacing:',
        txtPosition: 'Position:',
        txtBy: 'By:',
        txtNormal: 'Normal',
        txtExpanded: 'Expanded',
        txtCondensed: 'Condensed',
        txtRaised: 'Raised',
        txtLowered: 'Lowered',
        txtKerning: 'Kerning for fonts:',
        txtPointsAbove: 'Points and above',
        txtOpenType: 'OpenType Features',
        txtLigatures: 'Ligatures:',
        txtLigStandard: 'Standard Only',
        txtLigStandardContextual: 'Standard and Contextual',
        txtLigHistorical: 'Historical and Discretionary',
        txtLigAll: 'All',
        txtNumSpacing: 'Number spacing:',
        txtNumForms: 'Number forms:',
        txtStylisticSets: 'Stylistic sets:',
        txtDefault: 'Default',
        txtPreview: 'Preview'
    }, DE.Views.FontDialog || {}));
});

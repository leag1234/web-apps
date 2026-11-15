// SPDX-License-Identifier: AGPL-3.0-only
/**
 * SortDialog.js: the Sort dialog (up to three keys, type, order, header row) and its Options dialog
 * (field separator, case sensitivity).
 */
define([
], function () { 'use strict';

    DE.Views.SortOptionsDialog = Common.UI.Window.extend(_.extend({
        options: {
            width: 300,
            header: true,
            cls: 'modal-dlg',
            buttons: ['ok', 'cancel']
        },

        initialize: function (options) {
            _.extend(this.options, {title: this.textTitle}, options || {});
            var paragraphs = this.options.mode === 'paragraphs';
            this.template = [
                '<div class="box">',
                    paragraphs ? [
                        '<div style="margin-bottom: 8px;"><label class="font-weight-bold">' + this.textSeparator + '</label></div>',
                        '<div id="id-sort-opt-tabs" style="margin-bottom: 8px;"></div>',
                        '<div id="id-sort-opt-commas" style="margin-bottom: 5px;"></div>',
                        '<div style="margin-bottom: 16px;">',
                            '<div id="id-sort-opt-other" class="margin-right-10" style="display: inline-block;vertical-align: middle;"></div>',
                            '<div id="id-sort-opt-other-txt" style="display: inline-block;vertical-align: middle;"></div>',
                        '</div>'
                    ].join('') : '',
                    '<div id="id-sort-opt-case"></div>',
                '</div>'
            ].join('');
            this.options.tpl = _.template(this.template)(this.options);
            Common.UI.Window.prototype.initialize.call(this, this.options);
        },

        render: function () {
            Common.UI.Window.prototype.render.call(this);
            var $w = this.getChild(), s = this.options.settings || {}, paragraphs = this.options.mode === 'paragraphs';
            if (paragraphs) {
                this.rbTabs = new Common.UI.RadioBox({el: $w.find('#id-sort-opt-tabs'), labelText: this.textTabs, name: 'asc-radio-sort-separator', value: 'tab', checked: s.separator === 'tab'});
                this.rbCommas = new Common.UI.RadioBox({el: $w.find('#id-sort-opt-commas'), labelText: this.textCommas, name: 'asc-radio-sort-separator', value: 'comma', checked: s.separator === 'comma'});
                this.rbOther = new Common.UI.RadioBox({el: $w.find('#id-sort-opt-other'), labelText: this.textOther, name: 'asc-radio-sort-separator', value: 'other', checked: s.separator === 'other'});
                this.inputOther = new Common.UI.InputField({
                    el: $w.find('#id-sort-opt-other-txt'), style: 'width: 40px;', maxLength: 1,
                    validateOnChange: true, validateOnBlur: false, value: s.separatorChar || '', disabled: s.separator !== 'other'
                });
                var me = this;
                _.each([this.rbTabs, this.rbCommas, this.rbOther], function (rb) {
                    rb.on('change', function () { me.inputOther.setDisabled(!me.rbOther.getValue()); });
                });
            }
            this.chCase = new Common.UI.CheckBox({el: $w.find('#id-sort-opt-case'), labelText: this.textCase, value: !!s.caseSensitive});
            $w.find('.dlg-btn').on('click', _.bind(this.onBtnClick, this));
        },

        getFocusedComponents: function () {
            var list = this.options.mode === 'paragraphs' ? [this.rbTabs, this.rbCommas, this.rbOther, this.inputOther] : [];
            return list.concat([this.chCase]).concat(this.getFooterButtons());
        },

        getDefaultFocusableComponent: function () {
            return this.options.mode === 'paragraphs' ? this.rbTabs : this.chCase;
        },

        onBtnClick: function (event) {
            this.options.handler && this.options.handler.call(this, this, event.currentTarget.attributes['result'].value);
            this.close();
        },

        onPrimary: function () {
            this.options.handler && this.options.handler.call(this, this, 'ok');
            this.close();
            return false;
        },

        getSettings: function () {
            var res = {caseSensitive: this.chCase.getValue() === 'checked'};
            if (this.options.mode === 'paragraphs') {
                res.separator = this.rbCommas.getValue() ? 'comma' : (this.rbOther.getValue() ? 'other' : 'tab');
                res.separatorChar = res.separator === 'other' ? this.inputOther.getValue() : '';
            }
            return res;
        },

        textTitle: 'Sort Options',
        textSeparator: 'Separate fields at',
        textTabs: 'Tabs',
        textCommas: 'Commas',
        textOther: 'Other:',
        textCase: 'Case sensitive'
    }, DE.Views.SortOptionsDialog || {}));

    DE.Views.SortDialog = Common.UI.Window.extend(_.extend({
        options: {
            width: 460,
            header: true,
            cls: 'modal-dlg',
            buttons: ['ok', 'cancel']
        },

        initialize: function (options) {
            _.extend(this.options, {title: this.textTitle}, options || {});
            var info = this.options.info || {};
            this._mode = info.mode;
            this._settings = {separator: 'tab', separatorChar: '', caseSensitive: false};
            this._fields = info.fields || 1;
            this._columns = info.columns || 0;
            var rows = '';
            for (var i = 0; i < 3; i++) {
                rows += [
                    '<div class="sort-row" style="display: flex; align-items: center; margin-bottom: 8px;">',
                        '<label style="width: 62px; flex: none;">' + (i === 0 ? this.textSortBy : this.textThenBy) + '</label>',
                        '<div id="id-sort-field-' + i + '" style="width: 140px; margin-right: 8px;"></div>',
                        '<div id="id-sort-type-' + i + '" style="width: 100px; margin-right: 8px;"></div>',
                        '<div id="id-sort-order-' + i + '" style="width: 110px;"></div>',
                    '</div>'
                ].join('');
            }
            this.template = [
                '<div class="box">', rows,
                    '<div style="margin: 16px 0 6px;"><label class="font-weight-bold">' + this.textMyList + '</label></div>',
                    '<div id="id-sort-rb-header" style="margin-bottom: 6px;"></div>',
                    '<div id="id-sort-rb-noheader" style="margin-bottom: 16px;"></div>',
                    '<div id="id-sort-btn-options"></div>',
                '</div>'
            ].join('');
            this.options.tpl = _.template(this.template)(this.options);
            Common.UI.Window.prototype.initialize.call(this, this.options);
        },

        render: function () {
            Common.UI.Window.prototype.render.call(this);
            var me = this, $w = this.getChild(), info = this.options.info || {};

            this.cmbField = [];
            this.cmbType = [];
            this.cmbOrder = [];
            for (var i = 0; i < 3; i++) {
                this.cmbField.push(new Common.UI.ComboBox({
                    el: $w.find('#id-sort-field-' + i), cls: 'input-group-nr', style: 'width: 140px;', menuStyle: 'min-width: 140px; max-height: 200px;',
                    editable: false, takeFocusOnClose: true, data: this.fieldData(i)
                }));
                this.cmbField[i].setValue(i === 0 ? this.firstField() : -1);
                this.cmbType.push(new Common.UI.ComboBox({
                    el: $w.find('#id-sort-type-' + i), cls: 'input-group-nr', style: 'width: 100px;', menuStyle: 'min-width: 100px;',
                    editable: false, takeFocusOnClose: true,
                    data: [{value: 0, displayValue: this.textText}, {value: 1, displayValue: this.textNumber}, {value: 2, displayValue: this.textDate}]
                }));
                this.cmbType[i].setValue(0);
                this.cmbOrder.push(new Common.UI.ComboBox({
                    el: $w.find('#id-sort-order-' + i), cls: 'input-group-nr', style: 'width: 110px;', menuStyle: 'min-width: 110px;',
                    editable: false, takeFocusOnClose: true,
                    data: [{value: 1, displayValue: this.textAscending}, {value: 0, displayValue: this.textDescending}]
                }));
                this.cmbOrder[i].setValue(1);
            }

            this.rbHeader = new Common.UI.RadioBox({el: $w.find('#id-sort-rb-header'), labelText: this.textHeaderRow, name: 'asc-radio-sort-header', checked: !!info.header});
            this.rbNoHeader = new Common.UI.RadioBox({el: $w.find('#id-sort-rb-noheader'), labelText: this.textNoHeaderRow, name: 'asc-radio-sort-header', checked: !info.header});

            this.btnOptions = new Common.UI.Button({el: $w.find('#id-sort-btn-options'), caption: this.textOptions});
            this.btnOptions.on('click', _.bind(this.onOptions, this));

            $w.find('.dlg-btn').on('click', _.bind(this.onBtnClick, this));
        },

        // the whole paragraph (paragraphs) or the first column (tables)
        firstField: function () {
            return this._mode === 'table' ? 1 : 0;
        },

        fieldData: function (row) {
            var data = [], i;
            if (row > 0) data.push({value: -1, displayValue: this.textNone});
            if (this._mode === 'table') {
                for (i = 1; i <= this._columns; i++) data.push({value: i, displayValue: Common.Utils.String.format(this.textColumn, i)});
            } else {
                data.push({value: 0, displayValue: this.textParagraphs});
                if (this._fields > 1 || this._settings.separator !== 'tab')
                    for (i = 1; i <= this._fields; i++) data.push({value: i, displayValue: Common.Utils.String.format(this.textField, i)});
            }
            return data;
        },

        refreshFields: function () {
            for (var i = 0; i < 3; i++) {
                var cmb = this.cmbField[i], old = cmb.getValue();
                cmb.setData(this.fieldData(i));
                var valid = _.some(cmb.store ? cmb.store.models : [], function (m) { return m.get('value') === old; });
                cmb.setValue(valid ? old : (i === 0 ? this.firstField() : -1));
            }
        },

        onOptions: function () {
            var me = this;
            (new DE.Views.SortOptionsDialog({
                mode: this._mode,
                settings: this._settings,
                handler: function (dlg, result) {
                    if (result !== 'ok') return;
                    var s = dlg.getSettings();
                    _.extend(me._settings, s);
                    if (me._mode === 'paragraphs') {
                        var api = me.options.api;
                        var info = api && api.asc_GetSortInfo({separator: s.separator, separatorChar: s.separatorChar});
                        me._fields = (info && info.fields) || 1;
                        me.refreshFields();
                    }
                }
            })).show();
        },

        getFocusedComponents: function () {
            return this.cmbField.concat(this.cmbType, this.cmbOrder, [this.rbHeader, this.rbNoHeader, this.btnOptions]).concat(this.getFooterButtons());
        },

        getDefaultFocusableComponent: function () {
            return this.cmbField[0];
        },

        onBtnClick: function (event) {
            this.options.handler && this.options.handler.call(this, this, event.currentTarget.attributes['result'].value);
            this.close();
        },

        onPrimary: function () {
            this.options.handler && this.options.handler.call(this, this, 'ok');
            this.close();
            return false;
        },

        // settings for api.asc_Sort
        getSettings: function () {
            var keys = [];
            for (var i = 0; i < 3; i++) {
                var field = this.cmbField[i].getValue();
                if (field < 0) continue;
                keys.push({field: field, type: this.cmbType[i].getValue(), ascending: this.cmbOrder[i].getValue() == 1});
            }
            return {
                keys: keys,
                header: this.rbHeader.getValue(),
                caseSensitive: !!this._settings.caseSensitive,
                separator: this._settings.separator,
                separatorChar: this._settings.separatorChar,
                lang: this.options.lang
            };
        },

        textTitle: 'Sort',
        textSortBy: 'Sort by',
        textThenBy: 'Then by',
        textText: 'Text',
        textNumber: 'Number',
        textDate: 'Date',
        textAscending: 'Ascending',
        textDescending: 'Descending',
        textParagraphs: 'Paragraphs',
        textField: 'Field {0}',
        textColumn: 'Column {0}',
        textNone: '(none)',
        textMyList: 'My list has',
        textHeaderRow: 'Header row',
        textNoHeaderRow: 'No header row',
        textOptions: 'Options...'
    }, DE.Views.SortDialog || {}));
});

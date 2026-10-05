/*
 * SPDX-FileCopyrightText: 2026 ElixirOffice contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 *
 * Word-like style windows of the document editor:
 *  - DE.Views.StylesPane: the Styles pane (Ctrl+Alt+Shift+S, Styles group launcher), a floating
 *    non-modal window listing the styles, with the current one highlighted, a menu per style
 *    (Update to Match Selection, Modify, Delete/Revert) and New Style;
 *  - DE.Views.ApplyStylesDialog: Word's "Apply Styles" box (style name with completion, Reapply,
 *    Modify, Styles pane);
 *  - DE.Views.ModifyStyleDialog: Word's "Modify Style" dialog (name, based on, next style,
 *    formatting, preview, Add to the Styles gallery, Format > Font...).
 * The controller is DE.Controllers.WordFormatting.
 */

define([
    'common/main/lib/component/Window',
    'common/main/lib/component/ComboBox',
    'common/main/lib/component/ListView'
], function () { 'use strict';

    var enc = function (s) { return Common.Utils.String.htmlEncode(s || ''); };

    // ---------------------------------------------------------------------------------------------
    DE.Views.StylesPane = Common.UI.Window.extend(_.extend({
        options: {
            width: 260,
            height: 520,
            modal: false,
            resizable: true,
            minwidth: 220,
            minheight: 280,
            cls: 'modal-dlg word-styles-pane',
            id: 'window-word-styles-pane',
            buttons: null
        },

        initialize: function (options) {
            _.extend(this.options, {title: this.txtTitle}, options || {});
            this.options.tpl = [
                '<div class="box word-styles-pane-box">',
                    '<div class="word-styles-list"></div>',
                    '<div class="word-styles-options">',
                        '<div class="word-styles-chk-preview"></div>',
                    '</div>',
                    '<div class="word-styles-footer">',
                        '<div class="word-styles-btn-new"></div>',
                        '<div class="word-styles-btn-clear"></div>',
                    '</div>',
                '</div>'
            ].join('');
            this.handler = options.handler;
            Common.UI.Window.prototype.initialize.call(this, this.options);
        },

        render: function () {
            Common.UI.Window.prototype.render.call(this);
            var me = this, $win = this.getChild();

            this.list = new Common.UI.ListView({
                el: $win.find('.word-styles-list'),
                store: new Common.UI.DataViewStore(),
                tabindex: 1,
                cls: 'word-styles-listview',
                itemTemplate: _.template([
                    '<div id="<%= id %>" class="word-style-row<% if (isClear) { %> clear-all<% } %>" data-name="<%= Common.Utils.String.htmlEncode(name) %>">',
                        '<span class="word-style-name"<% if (preview) { %> style="<%= preview %>"<% } %>><%= Common.Utils.String.htmlEncode(caption) %></span>',
                        '<span class="word-style-type"><%= glyph %></span>',
                        '<% if (!isClear) { %><button type="button" class="word-style-menu-btn" tabindex="-1" aria-label="<%= Common.Utils.String.htmlEncode(caption) %>"><span class="caret"></span></button><% } %>',
                    '</div>'
                ].join(''))
            });
            this.list.on('item:click', function (view, itemView, record) {
                if (me._menuClick) { me._menuClick = false; return; }
                me.fire('apply', record);
            });
            this.list.on('entervalue', function () {
                var rec = me.list.getSelectedRec();
                rec && me.fire('apply', rec);
            });
            $win.find('.word-styles-list').on('mousedown', '.word-style-menu-btn', function (e) {
                me._menuClick = true;
                var name = $(e.currentTarget).closest('.word-style-row').attr('data-name');
                var rec = me.list.store.findWhere({name: name});
                rec && me.showStyleMenu(rec, e.currentTarget);
                e.preventDefault();
            });
            $win.find('.word-styles-list').on('contextmenu', '.word-style-row', function (e) {
                var name = $(e.currentTarget).attr('data-name');
                var rec = me.list.store.findWhere({name: name});
                rec && !rec.get('isClear') && me.showStyleMenu(rec, e.currentTarget);
                e.preventDefault();
                return false;
            });

            this.chPreview = new Common.UI.CheckBox({
                el: $win.find('.word-styles-chk-preview'),
                labelText: this.txtShowPreview,
                value: Common.localStorage.getBool('de-word-styles-preview', false)
            });
            this.chPreview.on('change', function (field, value) {
                Common.localStorage.setBool('de-word-styles-preview', value === 'checked' || value === true);
                me.fire('refresh');
            });

            this.btnNew = new Common.UI.Button({
                parentEl: $win.find('.word-styles-btn-new'),
                cls: 'btn-toolbar',
                iconCls: 'toolbar__icon btn-style-new',
                hint: this.txtNewStyle
            });
            this.btnNew.on('click', function () { me.fire('new'); });
            this.btnClear = new Common.UI.Button({
                parentEl: $win.find('.word-styles-btn-clear'),
                cls: 'btn-toolbar',
                iconCls: 'toolbar__icon btn-clearstyle',
                hint: this.txtClearAll
            });
            this.btnClear.on('click', function () { me.fire('clear'); });

            this.styleMenu = new Common.UI.Menu({
                menuAlign: 'tl-bl',
                items: [
                    {caption: this.txtUpdate, value: 'update'},
                    {caption: this.txtModify, value: 'modify'},
                    {caption: '--'},
                    {caption: this.txtDelete, value: 'delete'},
                    {caption: this.txtRevert, value: 'revert'}
                ]
            });
            this.styleMenu.on('item:click', function (menu, item) {
                me.fire(item.value, menu.record);
            });

            this.on('resize', function () { me.updateListHeight(); });
        },

        show: function () {
            Common.UI.Window.prototype.show.apply(this, arguments);
            this.positionAtRight();
            this.updateListHeight();
        },

        positionAtRight: function () {
            var w = this.getWidth ? this.getWidth() : this.options.width,
                left = Math.max(0, Common.Utils.innerWidth() - w - 48),
                top = Math.max(0, Math.min(150, Common.Utils.innerHeight() - this.options.height - 40));
            this.setPosition(left, top);
        },

        updateListHeight: function () {
            var $win = this.getChild(),
                h = $win.find('.body').height() - $win.find('.word-styles-options').outerHeight(true) - $win.find('.word-styles-footer').outerHeight(true) - 16;
            $win.find('.word-styles-list').css('height', Math.max(80, h));
            this.list.scroller && this.list.scroller.update({alwaysVisibleY: true});
        },

        fire: function (action, record) {
            this.handler && this.handler.call(this, action, record);
        },

        isPreview: function () {
            return this.chPreview && this.chPreview.getValue() === 'checked';
        },

        // styles: [{name, caption, type: 'p'|'c'|'l', isDefault, preview (css)}]
        setStyles: function (styles, current) {
            var arr = [{id: Common.UI.getId(), name: '', caption: this.txtClearAll, glyph: '', isClear: true, preview: ''}].concat(_.map(styles, function (s) {
                return _.extend({id: Common.UI.getId(), glyph: s.type === 'c' ? 'a' : (s.type === 'l' ? '¶a' : '¶'), isClear: false, preview: ''}, s);
            }));
            this.list.store.reset(arr);
            this.setCurrent(current);
        },

        setCurrent: function (name) {
            this._current = name;
            if (!this.list) return;
            var rec = name ? this.list.store.findWhere({name: name}) : null;
            if (rec) {
                this.list.selectRecord(rec, true);
                this.list.scrollToRecord(rec);
            } else
                this.list.deselectAll(true);
        },

        showStyleMenu: function (record, anchor) {
            var menu = this.styleMenu, me = this;
            menu.record = record;
            menu.items[0].setCaption(Common.Utils.String.format(this.txtUpdateTo, record.get('caption')));
            menu.items[3].setCaption(Common.Utils.String.format(this.txtDeleteNamed, record.get('caption')));
            menu.items[3].setVisible(!record.get('isDefault'));
            menu.items[4].setVisible(!!record.get('isDefault'));
            menu.items[4].setDisabled(!record.get('isChanged'));
            var $c = this.getChild().find('.word-styles-menu-container');
            if (!$c.length) {
                $c = $('<div class="word-styles-menu-container" style="position: absolute; z-index: 10000;"><div class="dropdown-toggle" data-toggle="dropdown"></div></div>');
                this.getChild().append($c);
                menu.render($c);
            }
            var off = $(anchor).offset(), woff = this.getChild().offset();
            $c.css({left: off.left - woff.left, top: off.top - woff.top + $(anchor).outerHeight()});
            setTimeout(function () { menu.show(); }, 0);
        },

        getFocusedComponents: function () {
            return [this.list, this.chPreview, this.btnNew, this.btnClear];
        },

        getDefaultFocusableComponent: function () {
            return this.list;
        },

        txtTitle: 'Styles',
        txtClearAll: 'Clear All',
        txtShowPreview: 'Show Preview',
        txtNewStyle: 'New Style',
        txtUpdate: 'Update to Match Selection',
        txtUpdateTo: 'Update {0} to Match Selection',
        txtModify: 'Modify...',
        txtDelete: 'Delete...',
        txtDeleteNamed: 'Delete {0}...',
        txtRevert: 'Revert to Default'
    }, DE.Views.StylesPane || {}));

    // ---------------------------------------------------------------------------------------------
    DE.Views.ApplyStylesDialog = Common.UI.Window.extend(_.extend({
        options: {
            width: 280,
            modal: false,
            cls: 'modal-dlg word-apply-styles',
            id: 'window-word-apply-styles',
            buttons: null
        },

        initialize: function (options) {
            _.extend(this.options, {title: this.txtTitle}, options || {});
            this.options.tpl = [
                '<div class="box">',
                    '<label class="input-label">' + this.txtStyleName + '</label>',
                    '<div class="word-apply-style-name input-group-nr" style="width: 100%;"></div>',
                    '<div class="word-apply-buttons">',
                        '<button type="button" class="btn btn-text-default word-apply-reapply">' + this.txtReapply + '</button>',
                        '<button type="button" class="btn btn-text-default word-apply-modify">' + this.txtModify + '</button>',
                        '<div class="word-apply-pane"></div>',
                    '</div>',
                '</div>'
            ].join('');
            this.handler = options.handler;
            this.styles = options.styles || [];
            Common.UI.Window.prototype.initialize.call(this, this.options);
        },

        render: function () {
            Common.UI.Window.prototype.render.call(this);
            var me = this, $win = this.getChild();
            this.cmbStyle = new Common.UI.ComboBox({
                el: $win.find('.word-apply-style-name'),
                cls: 'input-group-nr',
                menuStyle: 'min-width: 100%; max-height: 260px;',
                editable: true,
                data: _.map(this.styles, function (s) { return {value: s.name, displayValue: s.caption}; })
            });
            this.cmbStyle.on('selected', function (combo, record) { me.fire('apply', record.value); });
            this.cmbStyle.on('changed:after', function (combo, record) {
                var name = me.resolve(record && record.value !== undefined ? record.value : combo.getRawValue());
                name && me.fire('apply', name);
            });
            this.cmbStyle._input && this.cmbStyle._input.on('keyup', function (e) {
                if (e.keyCode === Common.UI.Keys.RETURN) {
                    var name = me.resolve(me.cmbStyle.getRawValue());
                    name && me.fire('apply', name);
                }
            });
            $win.find('.word-apply-reapply').on('click', function () {
                var name = me.resolve(me.cmbStyle.getRawValue());
                name && me.fire('apply', name);
            });
            $win.find('.word-apply-modify').on('click', function () {
                var name = me.resolve(me.cmbStyle.getRawValue());
                name && me.fire('modify', name);
            });
            this.btnPane = new Common.UI.Button({
                parentEl: $win.find('.word-apply-pane'),
                cls: 'btn-toolbar',
                iconCls: 'toolbar__icon btn-styles-pane',
                hint: this.txtStylesPane
            });
            this.btnPane.on('click', function () { me.fire('pane'); });
        },

        // display name or style name -> style name
        resolve: function (text) {
            text = (text || '').trim().toLowerCase();
            if (!text) return null;
            var s = _.find(this.styles, function (s) { return s.caption.toLowerCase() === text || s.name.toLowerCase() === text; }) ||
                    _.find(this.styles, function (s) { return s.caption.toLowerCase().indexOf(text) === 0; });
            return s ? s.name : null;
        },

        setCurrent: function (name) {
            var s = _.findWhere(this.styles, {name: name});
            this.cmbStyle && this.cmbStyle.setValue(name, s ? s.caption : name);
        },

        fire: function (action, value) {
            this.handler && this.handler.call(this, action, value);
        },

        getFocusedComponents: function () {
            return [this.cmbStyle, this.btnPane];
        },

        getDefaultFocusableComponent: function () {
            return this.cmbStyle;
        },

        txtTitle: 'Apply Styles',
        txtStyleName: 'Style Name:',
        txtReapply: 'Reapply',
        txtModify: 'Modify...',
        txtStylesPane: 'Styles (Ctrl+Alt+Shift+S)'
    }, DE.Views.ApplyStylesDialog || {}));

    // ---------------------------------------------------------------------------------------------
    DE.Views.ModifyStyleDialog = Common.UI.Window.extend(_.extend({
        options: {
            width: 540,
            cls: 'modal-dlg word-modify-style',
            id: 'window-word-modify-style',
            buttons: ['ok', 'cancel']
        },

        initialize: function (options) {
            _.extend(this.options, {title: options.isNew ? this.txtNewTitle : this.txtTitle}, options || {});
            var t = this;
            this.options.tpl = [
                '<div class="box">',
                    '<div class="word-ms-section">' + t.txtProperties + '</div>',
                    '<table class="word-ms-props"><tbody>',
                        '<tr><td><label>' + t.txtName + '</label></td><td><div class="word-ms-name"></div></td></tr>',
                        '<tr><td><label>' + t.txtType + '</label></td><td><div class="word-ms-type input-group-nr"></div></td></tr>',
                        '<tr><td><label>' + t.txtBasedOn + '</label></td><td><div class="word-ms-basedon input-group-nr"></div></td></tr>',
                        '<tr><td><label>' + t.txtNext + '</label></td><td><div class="word-ms-next input-group-nr"></div></td></tr>',
                    '</tbody></table>',
                    '<div class="word-ms-section">' + t.txtFormatting + '</div>',
                    '<div class="word-ms-row">',
                        '<div class="word-ms-font"></div><div class="word-ms-size"></div>',
                        '<div class="word-ms-bold"></div><div class="word-ms-italic"></div><div class="word-ms-underline"></div>',
                        '<div class="word-ms-color"></div>',
                    '</div>',
                    '<div class="word-ms-row">',
                        '<div class="word-ms-al-left"></div><div class="word-ms-al-center"></div><div class="word-ms-al-right"></div><div class="word-ms-al-just"></div>',
                        '<div class="word-ms-sep"></div>',
                        '<div class="word-ms-ls-1"></div><div class="word-ms-ls-15"></div><div class="word-ms-ls-2"></div>',
                        '<div class="word-ms-sep"></div>',
                        '<div class="word-ms-sp-more"></div><div class="word-ms-sp-less"></div>',
                        '<div class="word-ms-sep"></div>',
                        '<div class="word-ms-ind-dec"></div><div class="word-ms-ind-inc"></div>',
                    '</div>',
                    '<div class="word-ms-preview"><div class="word-ms-preview-before">' + _.times(6, function () { return t.txtPrevParagraph; }).join(' ') + '</div>',
                        '<div class="word-ms-preview-text">' + _.times(8, function () { return t.txtSample; }).join(' ') + '</div>',
                        '<div class="word-ms-preview-after">' + _.times(6, function () { return t.txtNextParagraph; }).join(' ') + '</div></div>',
                    '<div class="word-ms-row word-ms-bottom">',
                        '<div class="word-ms-qformat"></div>',
                        '<div class="word-ms-onlydoc"></div>',
                        '<div class="word-ms-format"></div>',
                    '</div>',
                '</div>'
            ].join('');
            this.styleInfo = options.styleInfo || {};
            this.styleNames = options.styleNames || [];
            this.fontStore = options.fontStore;
            this.handler = options.handler;
            this.textChanges = {};
            this.paraChanges = {};
            Common.UI.Window.prototype.initialize.call(this, this.options);
        },

        render: function () {
            Common.UI.Window.prototype.render.call(this);
            var me = this, $win = this.getChild(), info = this.styleInfo,
                tp = info.TextPr || {}, pp = info.ParaPr || {};
            $win.find('.dlg-btn').on('click', _.bind(this.onBtnClick, this));

            this.inputName = new Common.UI.InputField({
                el: $win.find('.word-ms-name'), style: 'width: 100%;', allowBlank: false,
                value: info.Caption || info.Name || '', disabled: !!info.IsDefault
            });
            this.cmbType = new Common.UI.ComboBox({
                el: $win.find('.word-ms-type'), cls: 'input-group-nr', editable: false, disabled: true, menuStyle: 'min-width: 100%;',
                data: [{value: 1, displayValue: this.txtParagraph}, {value: 4, displayValue: this.txtCharacter}]
            });
            this.cmbType.setValue(info.Type === 4 ? 4 : 1);
            var names = [{value: '', displayValue: this.txtNoStyle}].concat(_.map(this.styleNames, function (s) { return {value: s.name, displayValue: s.caption}; }));
            this.cmbBasedOn = new Common.UI.ComboBox({
                el: $win.find('.word-ms-basedon'), cls: 'input-group-nr', editable: false, menuStyle: 'min-width: 100%; max-height: 220px;',
                data: _.filter(names, function (n) { return n.value !== info.Name; })
            });
            this.cmbBasedOn.setValue(info.BasedOn || '');
            this.cmbBasedOn.on('selected', function (c, r) { me.basedOn = r.value; });
            this.cmbNext = new Common.UI.ComboBox({
                el: $win.find('.word-ms-next'), cls: 'input-group-nr', editable: false, menuStyle: 'min-width: 100%; max-height: 220px;',
                data: _.filter(names, function (n) { return n.value !== ''; }), disabled: info.Type === 4
            });
            this.cmbNext.setValue(info.Next || info.Name);
            this.cmbNext.on('selected', function (c, r) { me.next = r.value; });

            this.cmbFont = new Common.UI.ComboBoxFonts({
                el: $win.find('.word-ms-font'), cls: 'input-group-nr', style: 'width: 160px;',
                menuCls: 'scrollable-menu', menuStyle: 'min-width: 100%; max-height: 270px;',
                store: new Common.Collections.Fonts(), recent: 0, hint: this.txtFont
            });
            this.cmbFont.on('selected', function (c, r) { me.setText('FontName', r.name); });
            var sizes = [8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72];
            this.cmbSize = new Common.UI.ComboBox({
                el: $win.find('.word-ms-size'), cls: 'input-group-nr', style: 'width: 60px;', menuStyle: 'min-width: 60px; max-height: 220px;',
                editable: true, data: _.map(sizes, function (v) { return {value: v, displayValue: '' + v}; })
            });
            this.cmbSize.on('selected', function (c, r) { me.setText('FontSize', r.value); });
            this.cmbSize.on('changed:after', function (c, r) {
                var v = parseFloat(String(r && r.value !== undefined ? r.value : c.getRawValue()).replace(',', '.'));
                v > 0 && v <= 1638 && me.setText('FontSize', Math.round(v * 2) / 2);
            });
            var tbtn = function (cls, icon, hint, toggle, group) {
                var b = new Common.UI.Button({parentEl: $win.find(cls), cls: 'btn-toolbar', iconCls: 'toolbar__icon ' + icon, hint: hint, enableToggle: !!toggle, toggleGroup: group, allowDepress: !group});
                return b;
            };
            this.btnBold = tbtn('.word-ms-bold', 'btn-bold', this.txtBold, true);
            this.btnItalic = tbtn('.word-ms-italic', 'btn-italic', this.txtItalic, true);
            this.btnUnderline = tbtn('.word-ms-underline', 'btn-underline', this.txtUnderline, true);
            this.btnBold.on('toggle', function (b, s) { me.setText('Bold', s); });
            this.btnItalic.on('toggle', function (b, s) { me.setText('Italic', s); });
            this.btnUnderline.on('toggle', function (b, s) { me.setText('Underline', s); });
            this.btnColor = new Common.UI.ButtonColored({
                parentEl: $win.find('.word-ms-color'), cls: 'btn-toolbar', iconCls: 'toolbar__icon btn-fontcolor',
                hint: this.txtColor, auto: true, color: '000000', menu: true
            });
            this.btnColor.setMenu();
            this.mnuColor = this.btnColor.getPicker();
            this.btnColor.on('color:select', function (b, color) { me.setText('Color', Common.Utils.ThemeColor.getRgbColor(color)); });
            this.btnColor.on('auto:select', function () { var c = new Asc.asc_CColor(); c.put_auto(true); me.setText('Color', c); });

            var al = [['.word-ms-al-left', 'btn-align-left', this.txtLeft, 1], ['.word-ms-al-center', 'btn-align-center', this.txtCenter, 2],
                      ['.word-ms-al-right', 'btn-align-right', this.txtRight, 0], ['.word-ms-al-just', 'btn-align-just', this.txtJustify, 3]];
            this.alignButtons = _.map(al, function (a) {
                var b = tbtn(a[0], a[1], a[2], true, 'word-ms-align');
                b.on('toggle', function (btn, s) { s && me.setPara('Jc', a[3]); });
                b.jc = a[3];
                return b;
            });
            var ls = [['.word-ms-ls-1', '1', 1], ['.word-ms-ls-15', '1.5', 1.5], ['.word-ms-ls-2', '2', 2]];
            this.lineButtons = _.map(ls, function (a) {
                var b = new Common.UI.Button({parentEl: $win.find(a[0]), cls: 'btn-toolbar word-ms-ls', caption: a[1], hint: me.txtLineSpacing + ' ' + a[1], enableToggle: true, toggleGroup: 'word-ms-ls', allowDepress: false});
                b.on('toggle', function (btn, s) { s && me.setPara('Line', a[2]); });
                b.line = a[2];
                return b;
            });
            this.btnSpMore = tbtn('.word-ms-sp-more', 'btn-para-space-more', this.txtSpaceMore);
            this.btnSpLess = tbtn('.word-ms-sp-less', 'btn-para-space-less', this.txtSpaceLess);
            this.btnSpMore.on('click', function () { me.setPara('Before', Math.min(558.8, (me.curBefore() || 0) + 12 * 25.4 / 72)); });
            this.btnSpLess.on('click', function () { me.setPara('Before', Math.max(0, (me.curBefore() || 0) - 12 * 25.4 / 72)); });
            this.btnIndDec = tbtn('.word-ms-ind-dec', 'btn-decoffset', this.txtIndentDec);
            this.btnIndInc = tbtn('.word-ms-ind-inc', 'btn-incoffset', this.txtIndentInc);
            this.btnIndDec.on('click', function () { me.setPara('IndLeft', Math.max(0, (me.curIndent() || 0) - 12.7)); });
            this.btnIndInc.on('click', function () { me.setPara('IndLeft', (me.curIndent() || 0) + 12.7); });

            this.chQFormat = new Common.UI.CheckBox({el: $win.find('.word-ms-qformat'), labelText: this.txtAddToGallery, value: info.QFormat !== false});
            this.chQFormat.on('change', function (f, v) { me.qformat = (v === 'checked' || v === true); });
            this.radioDoc = new Common.UI.RadioBox({el: $win.find('.word-ms-onlydoc'), name: 'word-ms-scope', labelText: this.txtOnlyDoc, checked: true});

            this.btnFormat = new Common.UI.Button({
                parentEl: $win.find('.word-ms-format'), cls: 'btn-text-menu-default', caption: this.txtFormat,
                menu: new Common.UI.Menu({items: [{caption: this.txtFontDlg, value: 'font'}, {caption: this.txtParagraphDlg, value: 'paragraph', disabled: true}]})
            });
            this.btnFormat.menu.on('item:click', function (menu, item) {
                item.value === 'font' && me.handler && me.handler.call(me, 'font', me.getTextPr());
            });

            // initial values
            this.textPr = _.clone(tp);
            this.paraPr = _.clone(pp);
            this.btnBold.toggle(!!tp.Bold, true);
            this.btnItalic.toggle(!!tp.Italic, true);
            this.btnUnderline.toggle(!!tp.Underline, true);
            this.cmbSize.setValue(tp.FontSize || '');
            _.each(this.alignButtons, function (b) { b.toggle(b.jc === pp.Jc, true); });
            _.each(this.lineButtons, function (b) { b.toggle(pp.LineRule === Asc.linerule_Auto && Math.abs((pp.Line || 0) - b.line) < 0.01, true); });
            this.setColorButton(tp.Color);
            if (info.Type === 4) {
                _.each(this.alignButtons.concat(this.lineButtons, [this.btnSpMore, this.btnSpLess, this.btnIndDec, this.btnIndInc]), function (b) { b.setDisabled(true); });
            }
            this.afterRender();
        },

        afterRender: function () {
            this.mnuColor.updateColors(Common.Utils.ThemeColor.getEffectColors(), Common.Utils.ThemeColor.getStandartColors());
            if (this.fontStore) {
                this.cmbFont.fillFonts(this.fontStore);
                var name = this.textPr.FontName;
                var rec = name ? this.fontStore.findWhere({name: name}) : null;
                rec ? this.cmbFont.selectRecord(rec) : this.cmbFont.setValue(name || '');
            }
            this.updatePreview();
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

        curBefore: function () { return this.paraPr.Before; },
        curIndent: function () { return this.paraPr.IndLeft; },

        setText: function (key, value) {
            this.textChanges[key] = value;
            this.textPr[key] = value;
            this.updatePreview();
        },

        setPara: function (key, value) {
            this.paraChanges[key] = value;
            this.paraPr[key] = value;
            if (key === 'Line') {
                this.paraChanges.LineRule = Asc.linerule_Auto;
                this.paraPr.LineRule = Asc.linerule_Auto;
            }
            this.updatePreview();
        },

        // merge values from the Font dialog
        applyTextProps: function (props) {
            var me = this;
            _.each(props, function (v, k) { if (v !== undefined) me.setText(k, v); });
            this.btnBold.toggle(!!this.textPr.Bold, true);
            this.btnItalic.toggle(!!this.textPr.Italic, true);
            this.btnUnderline.toggle(!!this.textPr.Underline, true);
            this.cmbSize.setValue(this.textPr.FontSize || '');
            this.textPr.FontName && this.cmbFont.setValue(this.textPr.FontName);
            this.setColorButton(this.textPr.Color);
        },

        getTextPr: function () { return _.clone(this.textPr); },

        updatePreview: function () {
            var tp = this.textPr, pp = this.paraPr, css = {};
            css['font-family'] = tp.FontName ? '"' + tp.FontName + '"' : '';
            css['font-size'] = tp.FontSize ? Math.min(tp.FontSize, 28) + 'pt' : '';
            css['font-weight'] = tp.Bold ? 'bold' : 'normal';
            css['font-style'] = tp.Italic ? 'italic' : 'normal';
            var deco = [];
            tp.Underline && deco.push('underline');
            (tp.Strikeout || tp.DStrikeout) && deco.push('line-through');
            css['text-decoration'] = deco.join(' ') || 'none';
            css['font-variant'] = tp.SmallCaps ? 'small-caps' : 'normal';
            css['text-transform'] = tp.Caps ? 'uppercase' : 'none';
            var c = tp.Color;
            css['color'] = (c && c.get_auto && !c.get_auto()) ? '#' + Common.Utils.ThemeColor.getHexColor(c.get_r(), c.get_g(), c.get_b()) : '';
            if (this.styleInfo.Type !== 4) {
                css['text-align'] = ({0: 'right', 1: 'left', 2: 'center', 3: 'justify'})[pp.Jc] || 'left';
                css['line-height'] = (pp.LineRule === Asc.linerule_Auto && pp.Line) ? String(Math.max(1, pp.Line) * 1.2) : '';
                css['margin-left'] = pp.IndLeft ? Math.min(pp.IndLeft, 60) + 'mm' : '';
            }
            this.getChild().find('.word-ms-preview-text').css(css);
        },

        getChanges: function () {
            var res = {}, name = this.inputName.getValue().trim();
            if (!this.styleInfo.IsDefault && name && name !== (this.styleInfo.Caption || this.styleInfo.Name)) res.Name = name;
            if (this.basedOn !== undefined) res.BasedOn = this.basedOn;
            if (this.next !== undefined) res.Next = this.next;
            if (this.qformat !== undefined) res.QFormat = this.qformat;
            if (!_.isEmpty(this.textChanges)) res.TextPr = this.textChanges;
            if (!_.isEmpty(this.paraChanges)) res.ParaPr = this.paraChanges;
            return res;
        },

        onBtnClick: function (event) {
            this._handleInput(event.currentTarget.attributes['result'].value);
        },

        onPrimary: function () {
            this._handleInput('ok');
            return false;
        },

        _handleInput: function (state) {
            if (state === 'ok' && !this.styleInfo.IsDefault && this.inputName.getValue().trim() === '') {
                this.inputName.focus();
                return;
            }
            this.handler && this.handler.call(this, state, state === 'ok' ? this.getChanges() : null);
            this.close();
        },

        getFocusedComponents: function () {
            return [this.inputName, this.cmbBasedOn, this.cmbNext, this.cmbFont, this.cmbSize, this.btnBold, this.btnItalic, this.btnUnderline, this.btnColor,
                this.chQFormat, this.btnFormat].concat(this.getFooterButtons());
        },

        getDefaultFocusableComponent: function () {
            return this.styleInfo.IsDefault ? this.cmbBasedOn : this.inputName;
        },

        txtTitle: 'Modify Style',
        txtNewTitle: 'Create New Style from Formatting',
        txtProperties: 'Properties',
        txtName: 'Name:',
        txtType: 'Style type:',
        txtBasedOn: 'Style based on:',
        txtNext: 'Style for following paragraph:',
        txtParagraph: 'Paragraph',
        txtCharacter: 'Character',
        txtNoStyle: '(no style)',
        txtFormatting: 'Formatting',
        txtFont: 'Font',
        txtBold: 'Bold',
        txtItalic: 'Italic',
        txtUnderline: 'Underline',
        txtColor: 'Font Color',
        txtLeft: 'Align Left',
        txtCenter: 'Center',
        txtRight: 'Align Right',
        txtJustify: 'Justify',
        txtLineSpacing: 'Line spacing',
        txtSpaceMore: 'Increase Paragraph Spacing',
        txtSpaceLess: 'Decrease Paragraph Spacing',
        txtIndentDec: 'Decrease Indent',
        txtIndentInc: 'Increase Indent',
        txtPrevParagraph: 'Previous Paragraph',
        txtSample: 'Sample Text',
        txtNextParagraph: 'Following Paragraph',
        txtAddToGallery: 'Add to the Styles gallery',
        txtOnlyDoc: 'Only in this document',
        txtFormat: 'Format',
        txtFontDlg: 'Font...',
        txtParagraphDlg: 'Paragraph...'
    }, DE.Views.ModifyStyleDialog || {}));
});

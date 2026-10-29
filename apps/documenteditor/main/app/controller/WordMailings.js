/*
 * SPDX-FileCopyrightText: 2026 ElixirOffice contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 *
 * DE.Controllers.WordMailings: the Mailings tab of Word (Alt, M) on the mail merge engine of sdkjs
 * (asc_StartMailMerge, asc_AddMailMergeField, asc_PreviewMailMergeResult, asc_DownloadAsMailMerge).
 *
 *   Start Mail Merge   Select Recipients: a CSV or XLSX file of this computer (first row: field names), a URL,
 *                      or a spreadsheet of the storage when the integration offers one; Edit Recipient List
 *   Write & Insert     Insert Merge Field (Alt+Shift+F), Highlight Merge Fields
 *   Preview Results    Preview Results (Alt+Shift+K), First, Previous, the record number, Next, Last
 *   Finish             Finish & Merge: Edit Individual Documents (Alt+Shift+N, one .docx with every record),
 *                      Merge to PDF (Alt+Shift+M, one PDF), One PDF per Record, Send Email Messages (when the
 *                      integration gives the addresses)
 *   Alt+Shift+E        edit the recipient list
 * At most 100 records per merge (limit of the engine).
 */

define([
    'core',
    'common/main/lib/component/Button',
    'common/main/lib/component/InputField'
], function () {
    'use strict';

    var MAX_RECORDS = 100;

    DE.Controllers.WordMailings = Backbone.Controller.extend(_.extend({
        initialize: function () {
            var me = this;
            this._state = {count: 0, fields: [], preview: false, index: 0, highlight: false, queue: null};
            this.addListeners({
                'Toolbar': {
                    'render:after': function (toolbar) { me.onToolbarRender(toolbar); }
                }
            });
            Common.NotificationCenter.on('app:ready', this.onAppReady.bind(this));
        },

        appOptions: function () {
            var main = this.getApplication().getController('Main');
            return (main && main.appOptions) || {};
        },

        onToolbarRender: function (toolbar) {
            var mode = this.appOptions();
            if (!mode.isEdit || mode.isEditMailMerge || mode.isEditDiagram || mode.isEditOle || !mode.canUseMailMerge || this.toolbar) return;
            this.toolbar = toolbar;
            var me = this,
                big = function (id) { return '<span class="btn-slot text x-huge" id="' + id + '"></span>'; },
                small = function (id) { return '<span class="btn-slot text" id="' + id + '"></span>'; };
            // the preview of the results masks the toolbar (editing:disable): this panel stays above the mask
            var $panel = $('<section class="panel" data-tab="mailings" role="tabpanel" aria-labelledby="mailings" style="position:relative;z-index:1;">' +
                '<div class="group no-mask" data-group="mm-start">' + big('elx-mm-recipients') + big('elx-mm-edit') + '</div>' +
                '<div class="separator long"></div>' +
                '<div class="group no-mask" data-group="mm-write">' + big('elx-mm-field') + '<div class="elset-column"><div class="elset">' + small('elx-mm-highlight') + '</div></div></div>' +
                '<div class="separator long"></div>' +
                '<div class="group no-mask" data-group="mm-preview">' + big('elx-mm-preview') +
                    '<div class="elset-column"><div class="elset">' + small('elx-mm-first') + small('elx-mm-prev') +
                    '<span class="btn-slot" id="elx-mm-index" style="width:44px;margin:0 2px;"></span>' + small('elx-mm-next') + small('elx-mm-last') + '</div>' +
                    '<div class="elset"><span class="font-size-normal" id="elx-mm-count" style="padding:0 4px;"></span></div></div></div>' +
                '<div class="separator long"></div>' +
                '<div class="group no-mask" data-group="mm-finish">' + big('elx-mm-finish') + '</div>' +
                '</section>');

            var lock = [Common.enumLock.lostConnect, Common.enumLock.disableOnStart];
            var mk = function (id, cfg) {
                var b = new Common.UI.Button(_.extend({lock: lock, dataHint: '1', dataHintDirection: 'bottom', dataHintOffset: 'small'}, cfg));
                b.render($panel.find('#' + id));
                me.controls.push(b);
                return b;
            };
            this.controls = [];
            var canStorage = mode.canRequestSelectSpreadsheet || mode.canRequestMailMergeRecipients || mode.fileChoiceUrl && mode.fileChoiceUrl.indexOf('{documentType}') > -1;
            this.btnRecipients = mk('elx-mm-recipients', {cls: 'btn-toolbar x-huge icon-top', iconCls: 'toolbar__icon btn-mailmerge', caption: this.txtSelectRecipients, hint: this.tipSelectRecipients,
                menu: new Common.UI.Menu({items: [
                    {caption: this.txtFromFile, value: 'file'},
                    {caption: this.txtFromUrl, value: 'url', visible: !mode.isElixirLocal},
                    {caption: this.txtFromStorage, value: 'storage', visible: !!canStorage && !mode.isElixirLocal}
                ]})});
            this.btnRecipients.menu.on('item:click', function (menu, item) { me.selectRecipients(item.value); });
            this.btnEdit = mk('elx-mm-edit', {cls: 'btn-toolbar x-huge icon-top', iconCls: 'toolbar__icon btn-edit-text', caption: this.txtEditList, hint: this.tipEditList});
            this.btnEdit.on('click', function () { me.editRecipients(); });
            this.btnField = mk('elx-mm-field', {cls: 'btn-toolbar x-huge icon-top', iconCls: 'toolbar__icon btn-quick-field', caption: this.txtInsertField, hint: this.tipInsertField,
                menu: new Common.UI.Menu({items: []})});
            this.btnField.menu.on('item:click', function (menu, item) { me.insertField(item.value); });
            this.btnField.menu.on('show:before', function () { me.fillFieldMenu(); });
            this.btnHighlight = mk('elx-mm-highlight', {cls: 'btn-toolbar', iconCls: 'toolbar__icon btn-highlight', caption: this.txtHighlight, hint: this.tipHighlight, enableToggle: true});
            this.btnHighlight.on('click', function (b) { me.setHighlight(b.pressed); });
            this.btnPreview = mk('elx-mm-preview', {cls: 'btn-toolbar x-huge icon-top', iconCls: 'toolbar__icon btn-ic-reviewview', caption: this.txtPreview, hint: this.tipPreview, enableToggle: true});
            this.btnPreview.on('click', function (b) { me.setPreview(b.pressed); });
            this.btnFirst = mk('elx-mm-first', {cls: 'btn-toolbar', iconCls: 'toolbar__icon btn-firstitem', hint: this.tipFirst});
            this.btnPrev = mk('elx-mm-prev', {cls: 'btn-toolbar', iconCls: 'toolbar__icon btn-previous-field', hint: this.tipPrev});
            this.btnNext = mk('elx-mm-next', {cls: 'btn-toolbar', iconCls: 'toolbar__icon btn-next-field', hint: this.tipNext});
            this.btnLast = mk('elx-mm-last', {cls: 'btn-toolbar', iconCls: 'toolbar__icon btn-lastitem', hint: this.tipLast});
            this.btnFirst.on('click', function () { me.goRecord(0); });
            this.btnPrev.on('click', function () { me.goRecord(me._state.index - 1); });
            this.btnNext.on('click', function () { me.goRecord(me._state.index + 1); });
            this.btnLast.on('click', function () { me.goRecord(me._state.count - 1); });
            this.txtIndex = new Common.UI.InputField({el: $panel.find('#elx-mm-index'), allowBlank: true, value: '', style: 'width:44px;', maskExp: /[0-9]/, dataHint: '1', dataHintDirection: 'bottom', dataHintOffset: 'small'});
            this.txtIndex.on('changed:after', function (input, value) { var n = parseInt(value, 10); !isNaN(n) && me.goRecord(n - 1); });
            this.btnFinish = mk('elx-mm-finish', {cls: 'btn-toolbar x-huge icon-top', iconCls: 'toolbar__icon btn-big-ic-review', caption: this.txtFinish, hint: this.tipFinish,
                menu: new Common.UI.Menu({items: [
                    {caption: this.txtEditDocuments, value: 'docx'},
                    {caption: this.txtMergePdf, value: 'pdf'},
                    {caption: this.txtPdfPerRecord, value: 'pdf-each'},
                    {caption: this.txtSendEmail, value: 'email', visible: !!mode.canSendEmailAddresses}
                ]})});
            this.btnFinish.menu.on('item:click', function (menu, item) { me.finish(item.value); });

            toolbar.addTab({action: 'mailings', caption: this.txtMailings, extcls: 'canedit', layoutname: 'toolbar-mailings', dataHintTitle: 'M'}, $panel);
            // not in toolbar.lockControls: the preview of the results disables the toolbar, not this tab
            Common.Utils.lockControls(Common.enumLock.disableOnStart, true, {array: this.controls});
        },

        onAppReady: function () {
            var me = this;
            this.api = this.getApplication().getController('Viewport').getApi();
            if (!this.api || !this.api.asc_StartMailMerge) return;
            this.controls && Common.Utils.lockControls(Common.enumLock.disableOnStart, false, {array: this.controls});
            if (this.toolbar) {
                // after References, as in Word
                var $tabs = this.toolbar.$tabs || $(), $li = $tabs.find('> a[data-tab=mailings]').parent(), $links = $tabs.find('> a[data-tab=links]').parent();
                $links.length && $li.insertAfter($links);
                this.toolbar.$tabs && (this.toolbar.$tabs = this.toolbar.$tabs.parent().find('> li'));
                this.toolbar.setVisible('mailings', Common.UI.LayoutManager.isElementVisible('toolbar-collaboration-mailmerge'));
            }
            this.api.asc_registerCallback('asc_onStartMailMerge', _.bind(this.onStartMailMerge, this));
            this.api.asc_registerCallback('asc_onPreviewMailMergeResult', _.bind(this.onPreviewResult, this));
            this.api.asc_registerCallback('asc_onEndPreviewMailMergeResult', _.bind(this.onEndPreview, this));
            this.api.asc_registerCallback('asc_onEndAction', _.bind(this.onEndAction, this));
            document.addEventListener('keydown', _.bind(this.onKey, this), true);
            var kt = this.getApplication().getController('WordKeyTips');
            if (kt && kt.commands && kt.tabs && this.toolbar) {
                kt.tabs.mailings = 'M';
                kt.commands.mailings = [['#elx-mm-recipients', 'S'], ['#elx-mm-edit', 'D'], ['#elx-mm-field', 'I'], ['#elx-mm-highlight', 'H'],
                    ['#elx-mm-preview', 'P'], ['#elx-mm-first', 'F'], ['#elx-mm-prev', 'R'], ['#elx-mm-index', 'E'], ['#elx-mm-next', 'X'], ['#elx-mm-last', 'L'],
                    ['#elx-mm-finish', 'F2']];
            }
            this.updateControls();
        },

        // Word's mail merge keys: Alt+Shift+K, N, M, E, F (no macOS equivalent in Word)
        onKey: function (e) {
            if (!this.api || !this.toolbar || !e.altKey || !e.shiftKey || e.ctrlKey || e.metaKey || Common.Utils.ModalWindow.isVisible()) return;
            var a = document.activeElement;
            if (a && /^(INPUT|TEXTAREA|SELECT)$/.test(a.tagName) && a.id !== 'area_id') return;
            var action = {75: 'preview', 78: 'docx', 77: 'pdf', 69: 'edit', 70: 'field'}[e.keyCode];
            if (!action) return;
            var main = this.getApplication().getController('WordViews');
            if (main && main._state && main._state.view === 'outline' && action !== 'field') return; // outline keys
            e.preventDefault();
            e.stopPropagation();
            if (!this._state.count && action !== 'edit') { this.hint(this.txtNoRecipients); return; }
            if (action === 'preview') this.setPreview(!this._state.preview);
            else if (action === 'edit') this._state.count ? this.editRecipients() : this.selectRecipients('file');
            else if (action === 'field') this.openFieldMenu();
            else this.finish(action);
        },

        hint: function (text) {
            var keys = this.getApplication().getController('WordEditKeys');
            keys && keys.showHint && keys.showHint(text, 3000);
        },

        selectRecipients: function (type) {
            var tb = this.getApplication().getController('Toolbar');
            if (tb && tb.onSelectRecepientsClick) tb.onSelectRecepientsClick(type);
            else if (type === 'file') this.api.asc_StartMailMerge();
        },

        editRecipients: function () {
            var rm = this.getApplication().getController('RightMenu'), mm = rm && rm.rightmenu && rm.rightmenu.mergeSettings;
            if (mm && mm.onEditData) mm.onEditData();
        },

        onStartMailMerge: function () {
            this._state.count = this.api.asc_GetReceptionsCount();
            this._state.fields = this.api.asc_GetMailMergeFieldsNameList() || [];
            this._state.index = 0;
            this.updateControls();
            this.hint(Common.Utils.String.format(this.txtRecipientsLoaded, this._state.count));
        },

        fillFieldMenu: function () {
            var menu = this.btnField.menu;
            menu.removeAll();
            _.each(this._state.fields, function (name) { menu.addItem({caption: name, value: name}); });
        },

        openFieldMenu: function () {
            var $btn = this.btnField.cmpEl;
            this.toolbar && this.toolbar.setTab && this.toolbar.setTab('mailings');
            if (!this.btnField.menu.isVisible()) $btn.find('button').first().trigger('click');
            var me = this;
            _.defer(function () { me.btnField.menu.cmpEl && me.btnField.menu.cmpEl.find('a').first().focus(); });
        },

        insertField: function (name) {
            if (!name) return;
            this.api.asc_AddMailMergeField(name);
            Common.NotificationCenter.trigger('edit:complete');
        },

        setHighlight: function (on) {
            this._state.highlight = !!on;
            this.api.asc_SetHighlightMailMergeFields(!!on);
            Common.NotificationCenter.trigger('edit:complete');
        },

        setPreview: function (on) {
            if (!this._state.count) return;
            if (on) this.api.asc_PreviewMailMergeResult(Math.min(this._state.index, this._state.count - 1));
            else this.api.asc_EndPreviewMailMergeResult();
            Common.NotificationCenter.trigger('edit:complete');
        },

        goRecord: function (index) {
            if (!this._state.count) return;
            index = Math.max(0, Math.min(this._state.count - 1, index));
            this._state.index = index;
            if (this._state.preview) this.api.asc_PreviewMailMergeResult(index);
            this.updateControls();
            Common.NotificationCenter.trigger('edit:complete');
        },

        onPreviewResult: function (index) {
            this._state.preview = true;
            this._state.index = index;
            this.updateControls();
        },

        onEndPreview: function () {
            this._state.preview = false;
            this.updateControls();
        },

        // Finish & Merge: downloads made by the conversion service
        finish: function (kind) {
            if (!this._state.count) { this.hint(this.txtNoRecipients); return; }
            var last = Math.min(this._state.count, MAX_RECORDS) - 1;
            if (kind === 'email') {
                var rm = this.getApplication().getController('RightMenu'), mm = rm && rm.rightmenu && rm.rightmenu.mergeSettings;
                mm && mm.onMergeClick && mm.onMergeClick();
                return;
            }
            if (kind === 'pdf-each') {
                this._state.queue = _.range(0, last + 1);
                this.nextInQueue();
                return;
            }
            this.download(kind === 'pdf' ? Asc.c_oAscFileType.PDF : Asc.c_oAscFileType.DOCX, 0, last);
        },

        download: function (type, from, to) {
            if (!this.api.asc_DownloadAsMailMerge(type, from, to, true)) {
                this._state.queue = null;
                Common.UI.warning({msg: this.txtMergeError});
            }
        },

        nextInQueue: function () {
            var q = this._state.queue;
            if (!q || !q.length) { this._state.queue = null; return; }
            var i = q.shift();
            this.download(Asc.c_oAscFileType.PDF, i, i);
        },

        onEndAction: function (type, id) {
            if (id === Asc.c_oAscAsyncAction.DownloadMerge && this._state.queue) {
                if (this.appOptions().isElixirLocal && this.api.eoLastMailMergeExport !== 'saved') {
                    this._state.queue = null;
                    return;
                }
                var me = this;
                _.defer(function () { me.nextInQueue(); });
            }
        },

        updateControls: function () {
            if (!this.btnField) return;
            var has = this._state.count > 0, st = this._state;
            this.btnEdit.setDisabled(!has);
            this.btnField.setDisabled(!has || st.preview);
            this.btnHighlight.setDisabled(!has);
            this.btnPreview.setDisabled(!has);
            this.btnFinish.setDisabled(!has);
            this.btnPreview.toggle(st.preview, true);
            this.btnFirst.setDisabled(!st.preview || st.index <= 0);
            this.btnPrev.setDisabled(!st.preview || st.index <= 0);
            this.btnNext.setDisabled(!st.preview || st.index >= st.count - 1);
            this.btnLast.setDisabled(!st.preview || st.index >= st.count - 1);
            this.txtIndex.setDisabled(!st.preview);
            this.txtIndex.setValue(has ? String(st.index + 1) : '');
            $('#elx-mm-count').text(has ? Common.Utils.String.format(this.txtRecordOf, st.count) : '');
        },

        getState: function () {
            return {count: this._state.count, fields: this._state.fields.slice(), preview: this._state.preview, index: this._state.index, highlight: this._state.highlight};
        },

        txtMailings: 'Mailings',
        txtSelectRecipients: 'Select Recipients',
        tipSelectRecipients: 'Choose the list of people you want to send the letter to (a CSV or Excel file whose first row has the field names)',
        txtFromFile: 'Use an Existing List...',
        txtFromUrl: 'From URL...',
        txtFromStorage: 'From Storage...',
        txtEditList: 'Edit Recipient List',
        tipEditList: 'Make changes to the list of recipients (Alt+Shift+E)',
        txtInsertField: 'Insert Merge Field',
        tipInsertField: 'Add a field from your recipient list to the document (Alt+Shift+F)',
        txtHighlight: 'Highlight Merge Fields',
        tipHighlight: 'Highlight the fields you have inserted into the document',
        txtPreview: 'Preview Results',
        tipPreview: 'Replace the merge fields with data from your recipient list (Alt+Shift+K)',
        tipFirst: 'First Record',
        tipPrev: 'Previous Record',
        tipNext: 'Next Record',
        tipLast: 'Last Record',
        txtRecordOf: 'of {0} records',
        txtFinish: 'Finish & Merge',
        tipFinish: 'Complete the mail merge: one document with every record, a PDF, or one PDF per record',
        txtEditDocuments: 'Edit Individual Documents... (.docx)',
        txtMergePdf: 'Merge to PDF',
        txtPdfPerRecord: 'One PDF per Record',
        txtSendEmail: 'Send Email Messages...',
        txtNoRecipients: 'Select the recipients first (Mailings > Select Recipients)',
        txtRecipientsLoaded: '{0} recipients',
        txtMergeError: 'The merge could not be made.'
    }, DE.Controllers.WordMailings || {}));
});

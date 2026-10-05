/*
 * SPDX-FileCopyrightText: 2026 ElixirOffice contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 *
 * ElixirOffice window chrome: Excel-like backstage (spreadsheet editor).
 * The backstage is shared (common/main/lib/elixir-chrome/Backstage.js); this
 * subclass gives Excel's file formats, workbook wording, the Protect Workbook
 * menu of Excel (Protect Current Sheet, Protect Workbook Structure) and no
 * tracked-changes row in Inspect Workbook. Translations: locale
 * SSE.Views.ElixirBackstage.*.
 */
define([
    'core',
    'common/main/lib/elixir-chrome/Backstage'
], function () {
    'use strict';

    var STRINGS = {
        txtBlankDocument: 'Blank workbook', txtMoreDocuments: 'More workbooks',
        txtNoRecent: 'Workbooks you open recently appear here.', txtNoMatch: 'No workbook matches your search.',
        txtSearchRecent: 'Search recent workbooks',
        txtShareWithPeopleDesc: 'Invite people and choose what they can do with this workbook.',
        txtCopyLinkDesc: 'Copy a link to this workbook to paste it in a message.',
        txtEmailLinkDesc: 'Send a link to this workbook from your email app.',
        txtSendCopyDesc: 'Download a copy of the workbook to attach it to a message.',
        txtSaveAsLead: 'Download a copy of this workbook in the format of your choice.',
        txtRenameDesc: 'Change the name of this workbook.',
        txtProtectDocument: 'Protect Workbook', txtProtectDocumentDesc: 'Control what types of changes people can make to this workbook.',
        txtInspectDocument: 'Inspect Workbook', txtInspectDocumentDesc: 'Before publishing this file, check whether it contains comments or personal information.',
        txtDocumentTypes: 'Workbook file types',
        txtProtectSheet: 'Protect Current Sheet', txtProtectSheetDesc: 'Control what types of changes people can make to the current sheet.',
        txtProtectStructure: 'Protect Workbook Structure', txtProtectStructureDesc: 'Prevent unwanted changes to the structure of the workbook, such as adding sheets.',
        txtFmtXlsx: 'Excel workbook (*.xlsx)', txtFmtXlsxDesc: 'The standard format, opens everywhere',
        txtFmtXlsm: 'Macro-enabled workbook (*.xlsm)', txtFmtXlsmDesc: 'Keeps the macros of the workbook',
        txtFmtXlsb: 'Binary workbook (*.xlsb)', txtFmtXlsbDesc: 'Fast to open and save for very large workbooks',
        txtFmtXltx: 'Excel template (*.xltx)', txtFmtXltxDesc: 'Starting point for new workbooks',
        txtFmtOds: 'OpenDocument spreadsheet (*.ods)', txtFmtOdsDesc: 'Open standard, for LibreOffice and others',
        txtFmtOts: 'OpenDocument template (*.ots)', txtFmtOtsDesc: 'Open standard template',
        txtFmtCsv: 'CSV UTF-8 (comma delimited) (*.csv)', txtFmtCsvDesc: 'Values of the current sheet only, as text',
        txtFmtPdfDescSheet: 'Fixed layout, easy to share and print'
    };

    SSE.Views.ElixirBackstage = Common.Views.ElixirBackstage.extend(_.extend({
        defaultFileType: 'xlsx',

        formats: function () {
            var T = Asc.c_oAscFileType;
            return [
                { type: T.XLSX, ext: '.xlsx', icon: 'table', name: 'txtFmtXlsx', desc: 'txtFmtXlsxDesc', group: 'doc' },
                { type: T.XLSM, ext: '.xlsm', icon: 'code', name: 'txtFmtXlsm', desc: 'txtFmtXlsmDesc', group: 'doc', onlyFor: 'xlsm' },
                { type: T.XLSB, ext: '.xlsb', icon: 'table', name: 'txtFmtXlsb', desc: 'txtFmtXlsbDesc', group: 'doc' },
                { type: T.XLTX, ext: '.xltx', icon: 'template', name: 'txtFmtXltx', desc: 'txtFmtXltxDesc', group: 'doc' },
                { type: T.ODS, ext: '.ods', icon: 'table', name: 'txtFmtOds', desc: 'txtFmtOdsDesc', group: 'doc' },
                { type: T.OTS, ext: '.ots', icon: 'template', name: 'txtFmtOts', desc: 'txtFmtOtsDesc', group: 'doc' },
                { type: T.CSV, ext: '.csv', icon: 'text', name: 'txtFmtCsv', desc: 'txtFmtCsvDesc', group: 'other' },
                { type: T.PDF, ext: '.pdf', icon: 'pdf', name: 'txtFmtPdf', desc: 'txtFmtPdfDesc', group: 'pdf' },
                { type: T.PDFA, ext: '.pdf', icon: 'pdf', name: 'txtFmtPdfa', desc: 'txtFmtPdfaDesc', group: 'pdf' },
                { type: T.PNG, ext: '.zip', icon: 'image', name: 'txtFmtPng', desc: 'txtFmtImgDesc', group: 'image' },
                { type: T.JPG, ext: '.zip', icon: 'image', name: 'txtFmtJpg', desc: 'txtFmtImgDesc', group: 'image' }
            ];
        },

        defaultFormat: function () { return Asc.c_oAscFileType.XLSX; },

        canRestrictEditing: function () {
            var mode = this.mode || {};
            return !!(this.ctrl('WBProtection') && mode.isEdit);
        },

        // spreadsheets have no tracked changes
        countTrackedChanges: function () { return null; },

        showProtectMenu: function ($btn, e) {
            var me = this, fm = this.fm, mode = this.mode || {};
            if (!this.protectMenu) {
                var items = [];
                if (fm.panels['protect'] && mode.isPasswordSupport)
                    items.push({ caption: this.txtEncryptPassword, value: 'password' });
                if (mode.isEdit && this.ctrl('WBProtection')) {
                    items.push({ caption: this.txtProtectSheet, value: 'sheet' });
                    items.push({ caption: this.txtProtectStructure, value: 'workbook' });
                }
                if (fm.panels['protect'] && mode.isSignatureSupport)
                    items.push({ caption: this.txtAddSignature, value: 'signature' });
                if (!items.length) return;
                this.protectMenu = new Common.UI.Menu({ cls: 'shifted-right', menuAlign: 'tl-bl', items: items });
                this.protectMenu.on('item:click', function (menu, item) {
                    if (item.value === 'password' || item.value === 'signature') { me.select('protect'); return; }
                    fm.hide();
                    var prot = me.ctrl('WBProtection');
                    if (!prot) return;
                    item.value === 'sheet' ? prot.onSheetClick(true) : prot.onWorkbookClick(true);
                });
            }
            var $holder = this.protectHolder || (this.protectHolder = $('<div class="elx-menu-holder" style="position:absolute;"></div>').appendTo(fm.$el));
            var off = $btn.offset(), base = fm.$el.offset();
            $holder.css({ left: off.left - base.left, top: off.top - base.top + $btn.outerHeight(), width: 1, height: 1 });
            if (!this.protectMenu.rendered) this.protectMenu.render($holder);
            this.protectMenu.show();
            e && e.stopPropagation();
        }
    }, STRINGS, SSE.Views.ElixirBackstage || {}));
});

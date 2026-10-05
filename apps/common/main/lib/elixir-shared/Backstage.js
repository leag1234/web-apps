/*
 * ElixirOffice window chrome: Word-like backstage (File menu).
 *
 * Builds on DE.Views.FileMenu without changing it: the pane gets Word's order
 * (Home, New, Open, Info, Save a Copy, Save As, Print, Share, Export, Close,
 * Options at the bottom), new pages (Home, Open, Share, Export), a Word-like
 * Info page (action tiles + properties column), Save As as a list of formats
 * and the print preview page in the browser too.
 *
 * Shared copy for all editors (ElixirOffice, apps/common/main/lib/elixir-shared/):
 * generalized from documenteditor/main/app/elixir-chrome/Backstage.js.
 * The host editor sets window.ElixirShared.app (DE, PE, SSE) and options
 * before creating the views (see ElixirShared.init in Shared.js).
 *
 * Licensed under the GNU AGPL v3 like the rest of web-apps.
 */
define([
    'core',
    'common/main/lib/elixir-shared/Shared',
    'common/main/lib/elixir-shared/Icons'
], function () {
    'use strict';

    var Icons = window.ElixirChromeIcons;
    var esc = function (s) { return Common.Utils.String.htmlEncode(s === undefined || s === null ? '' : String(s)); };

    // Document formats: [type, ext, icon, name key, description key]
    // the host editor can give its own list (ElixirShared.opts.formats)
    var FORMATS = function () {
        var custom = ElixirShared.opt('formats');
        if (custom) return custom();
        var T = Asc.c_oAscFileType;
        return [
            { type: T.DOCX, ext: '.docx', icon: 'doc', name: 'txtFmtDocx', desc: 'txtFmtDocxDesc', group: 'doc' },
            { type: T.DOCM, ext: '.docm', icon: 'code', name: 'txtFmtDocm', desc: 'txtFmtDocmDesc', group: 'doc', onlyFor: 'docm' },
            { type: T.DOTX, ext: '.dotx', icon: 'template', name: 'txtFmtDotx', desc: 'txtFmtDotxDesc', group: 'doc' },
            { type: T.ODT, ext: '.odt', icon: 'doc', name: 'txtFmtOdt', desc: 'txtFmtOdtDesc', group: 'doc' },
            { type: T.OTT, ext: '.ott', icon: 'template', name: 'txtFmtOtt', desc: 'txtFmtOttDesc', group: 'doc' },
            { type: T.RTF, ext: '.rtf', icon: 'doc', name: 'txtFmtRtf', desc: 'txtFmtRtfDesc', group: 'doc' },
            { type: T.TXT, ext: '.txt', icon: 'text', name: 'txtFmtTxt', desc: 'txtFmtTxtDesc', group: 'doc' },
            { type: T.PDF, ext: '.pdf', icon: 'pdf', name: 'txtFmtPdf', desc: 'txtFmtPdfDesc', group: 'pdf' },
            { type: T.PDFA, ext: '.pdf', icon: 'pdf', name: 'txtFmtPdfa', desc: 'txtFmtPdfaDesc', group: 'pdf' },
            { type: T.MD, ext: '.md', icon: 'markdown', name: 'txtFmtMd', desc: 'txtFmtMdDesc', group: 'other' },
            { type: T.HTML, ext: '.html', icon: 'globe', name: 'txtFmtHtml', desc: 'txtFmtHtmlDesc', group: 'other' },
            { type: T.EPUB, ext: '.epub', icon: 'book', name: 'txtFmtEpub', desc: 'txtFmtEpubDesc', group: 'other' },
            { type: T.FB2, ext: '.fb2', icon: 'book', name: 'txtFmtFb2', desc: 'txtFmtFb2Desc', group: 'other' },
            { type: T.PNG, ext: '.zip', icon: 'image', name: 'txtFmtPng', desc: 'txtFmtImgDesc', group: 'image' },
            { type: T.JPG, ext: '.zip', icon: 'image', name: 'txtFmtJpg', desc: 'txtFmtImgDesc', group: 'image' }
        ];
    };

    // English strings; translations: locale Common.Views.ElixirBackstage.*
    var STRINGS = {
        txtHome: 'Home', txtNew: 'New', txtOpen: 'Open', txtInfo: 'Info', txtSaveCopy: 'Save a Copy', txtSaveAs: 'Save As',
        txtDownload: 'Download', txtPrint: 'Print', txtShare: 'Share', txtExport: 'Export', txtClose: 'Close', txtOptions: 'Options',
        txtHelp: 'Help', txtBack: 'Back', txtAbout: 'About',
        txtGoodMorning: 'Good morning', txtGoodAfternoon: 'Good afternoon', txtGoodEvening: 'Good evening',
        txtMoreTemplates: 'More templates', txtMoreDocuments: 'More documents', txtRecent: 'Recent', txtName: 'Name', txtModified: 'Date modified',
        txtNoRecent: 'Documents you open recently appear here.', txtNoMatch: 'No document matches your search.',
        txtToday: 'Today, {0}', txtYesterday: 'Yesterday, {0}', txtBlankDocument: 'Blank document',
        txtBrowse: 'Browse', txtBrowseDesc: 'Open a file from this device', txtMyFiles: 'My files', txtStartPageDesc: 'Open the start page in a new tab',
        txtSearchRecent: 'Search recent documents', txtUploading: 'Uploading {0}…', txtUploadError: 'The file could not be uploaded.',
        txtShareWithPeople: 'Share with people', txtShareWithPeopleDesc: 'Invite people and choose what they can do with this document.',
        txtCopyLink: 'Copy link', txtCopyLinkDesc: 'Copy a link to this document to paste it in a message.', txtLinkCopied: 'Link copied',
        txtEmailLink: 'Email a link', txtEmailLinkDesc: 'Send a link to this document from your email app.',
        txtSendCopy: 'Send a copy', txtSendCopyDesc: 'Download a copy of the document to attach it to a message.',
        txtPeopleWithAccess: 'People with access',
        txtCreatePdf: 'Create PDF document', txtChangeType: 'Change file type',
        txtPdfBullet1: 'Preserves layout, formatting, fonts and images', txtPdfBullet2: 'Content cannot be easily changed',
        txtPdfBullet3: 'Free viewers are available on every device', txtCreatePdfBtn: 'Create PDF', txtCreatePdfaBtn: 'Create PDF/A (archiving)',
        txtDocumentTypes: 'Document file types', txtOtherTypes: 'Other file types', txtImageTypes: 'Images',
        txtSaveAsLead: 'Download a copy of this document in the format of your choice.', txtDownloadCopy: 'Download a copy',
        txtRename: 'Rename', txtRenameDesc: 'Change the name of this document.',
        txtProperties: 'Properties', txtProtectDocument: 'Protect Document', txtProtectDocumentDesc: 'Control what types of changes people can make to this document.',
        txtInspectDocument: 'Inspect Document', txtInspectDocumentDesc: 'Before publishing this file, check whether it contains comments, tracked changes or personal information.',
        txtVersionHistory: 'Version History', txtVersionHistoryDesc: 'View and restore previous versions.', txtVersionHistoryNone: 'Version history is not available for this document.',
        txtEncryptPassword: 'Encrypt with password', txtRestrictEditing: 'Restrict editing', txtAddSignature: 'Add a digital signature',
        txtInspComments: 'Comments', txtInspChanges: 'Tracked changes', txtInspProps: 'Document properties and personal information',
        txtInspCommentsFound: '{0} comment(s) found.', txtInspChangesFound: '{0} tracked change(s) found.', txtInspAuthor: 'Author: {0}',
        txtInspNone: 'Nothing found.', txtRemoveAll: 'Remove all', txtRemoved: 'Removed',
        txtBrowserPrinter: 'Printer', txtBrowserPrinterDesc: 'Choose the printer, copies and two-sided printing in the print dialog of your browser.',
        txtFmtDocx: 'Word document (*.docx)', txtFmtDocxDesc: 'The standard format, opens everywhere',
        txtFmtDocm: 'Macro-enabled document (*.docm)', txtFmtDocmDesc: 'Keeps the macros of the document',
        txtFmtDotx: 'Document template (*.dotx)', txtFmtDotxDesc: 'Starting point for new documents',
        txtFmtOdt: 'OpenDocument text (*.odt)', txtFmtOdtDesc: 'Open standard, for LibreOffice and others',
        txtFmtOtt: 'OpenDocument template (*.ott)', txtFmtOttDesc: 'Open standard template',
        txtFmtRtf: 'Rich Text Format (*.rtf)', txtFmtRtfDesc: 'Simple formatting, very compatible',
        txtFmtTxt: 'Plain text (*.txt)', txtFmtTxtDesc: 'Text only, no formatting',
        txtFmtPdf: 'PDF (*.pdf)', txtFmtPdfDesc: 'Fixed layout, easy to share and print',
        txtFmtPdfa: 'PDF/A (*.pdf)', txtFmtPdfaDesc: 'PDF for long-term archiving',
        txtFmtMd: 'Markdown (*.md)', txtFmtMdDesc: 'Lightweight text format for the web and wikis',
        txtFmtHtml: 'Web page (*.html, zipped)', txtFmtHtmlDesc: 'For publishing on a website',
        txtFmtEpub: 'EPUB e-book (*.epub)', txtFmtEpubDesc: 'For e-readers',
        txtFmtFb2: 'FictionBook (*.fb2)', txtFmtFb2Desc: 'E-book format',
        txtBlankPresentation: 'Blank presentation', txtBlankWorkbook: 'Blank workbook',
        txtPresentationTypes: 'Presentation file types',
        txtFmtPptx: 'PowerPoint presentation (*.pptx)', txtFmtPptxDesc: 'The standard format, opens everywhere',
        txtFmtPptm: 'Macro-enabled presentation (*.pptm)', txtFmtPptmDesc: 'Keeps the macros of the presentation',
        txtFmtPotx: 'Presentation template (*.potx)', txtFmtPotxDesc: 'Starting point for new presentations',
        txtFmtPpsx: 'PowerPoint show (*.ppsx)', txtFmtPpsxDesc: 'Opens directly as a slide show',
        txtFmtOdp: 'OpenDocument presentation (*.odp)', txtFmtOdpDesc: 'Open standard, for LibreOffice and others',
        txtFmtOtp: 'OpenDocument presentation template (*.otp)', txtFmtOtpDesc: 'Open standard template',
        txtFmtImgSlideDesc: 'One image per slide, in a zip file',
        txtPdfSlideBullet1: 'Preserves the layout, fonts and images of every slide',
        txtFmtPng: 'PNG images (*.png)', txtFmtJpg: 'JPEG images (*.jpg)', txtFmtImgDesc: 'One image per page, in a zip file'
    };

    // A file-menu page: a content box in the backstage that FileMenu shows/hides.
    var Page = Common.UI.BaseView.extend({
        initialize: function (options) {
            Common.UI.BaseView.prototype.initialize.call(this, arguments);
            this.bs = options.backstage;
            this.menu = options.menu;
        },
        show: function () {
            this.$el.show();
            this.refresh && this.refresh.apply(this, arguments);
            this.scroller && this.scroller.update();
        },
        hide: function () { this.$el.hide(); },
        isVisible: function () { return this.$el.is(':visible'); },
        setScroller: function () {
            if (!this.scroller)
                this.scroller = new Common.UI.Scroller({ el: this.$el, suppressScrollX: true, alwaysVisibleY: true });
        }
    });

    Common.Views.ElixirBackstage = Backbone.View.extend(_.extend({

        initialize: function (options) {
            this.controller = options.controller;
            this.fm = options.menu;
            this.mode = {};
            this.patchFileMenu();
        },

        L: function (key) { return this[key] || key; },

        icon: function (name, size, cls) { return Icons ? Icons.svg(name, size, cls) : ''; },

        // ---- FileMenu integration ---------------------------------------------------------

        patchFileMenu: function () {
            var me = this, fm = this.fm;
            var show = fm.show, applyMode = fm.applyMode;

            // Word opens the backstage on Home
            fm.show = function (panel, opts) {
                if (!this.rendered && this.mode && Common.Controllers.LaunchController.isScriptLoaded()) this.render();
                if (panel === undefined && !this.isVisible() && me.pages && me.pages['elx-home'] && me.isItemVisible('elx-home')) {
                    this.active = undefined;
                    panel = 'elx-home';
                }
                return show.call(this, panel, opts);
            };
            fm.applyMode = function () {
                var res = applyMode.apply(this, arguments);
                me.afterApplyMode();
                return res;
            };
        },

        isItemVisible: function (action) {
            var item = (this.items && this.items[action]) || (this.fm && this.fm._getMenuItem && this.fm._getMenuItem(action));
            return !!(item && item.$el && item.$el.css('display') !== 'none');
        },

        // called on FileMenu 'render:after' (the first applyMode already ran)
        onRender: function () {
            var fm = this.fm, me = this,
                $menu = fm.$el.find('.panel-menu'),
                $ctx = fm.$el.find('.panel-context');
            if (!$menu.length || this.rendered) return;
            this.rendered = true;
            fm.$el.addClass('elx-backstage');

            var addItem = function (id, action, caption, icon, where, $ref) {
                var $li = $('<li id="' + id + '" class="fm-btn"></li>');
                $ref && $ref.length ? $li[where]($ref) : $menu.append($li);
                var item = new Common.UI.MenuItem({
                    el: $li,
                    action: action,
                    caption: caption,
                    canFocused: false,
                    dataHint: 1,
                    dataHintDirection: 'left-top',
                    dataHintOffset: [-2, 22],
                    iconCls: 'menu__icon'
                });
                fm.items.push(item);
                return item;
            };
            // Office order in editors whose template differs from the Word one (no-op there):
            // Info in the first block, Options last
            var $recent = $menu.find('#fm-btn-recent'), $info = $menu.find('#fm-btn-info');
            if ($recent.length && $info.length && $info.prevAll('#fm-btn-save').length) $info.insertAfter($recent);
            var $settings = $menu.find('#fm-btn-settings');
            $settings.length && $settings.nextAll('#fm-btn-help, #fm-btn-suggest').length && $settings.insertAfter($menu.find('#fm-btn-help, #fm-btn-suggest').last());
            this.items = {};
            this.items['elx-home'] = addItem('fm-btn-elx-home', 'elx-home', this.txtHome, 'home', 'insertBefore', $menu.find('#fm-btn-create'));
            this.items['elx-open'] = addItem('fm-btn-elx-open', 'elx-open', this.txtOpen, 'open', 'insertAfter', $menu.find('#fm-btn-create'));
            // Save As (download formats) right after Save a Copy, then Print
            $menu.find('#fm-btn-download').insertAfter($menu.find('#fm-btn-save-desktop'));
            var $afterPrint = $menu.find('#fm-btn-print-with-preview');
            this.items['elx-share'] = addItem('fm-btn-elx-share', 'elx-share', this.txtShare, 'share', 'insertAfter', $afterPrint);
            this.items['elx-export'] = addItem('fm-btn-elx-export', 'elx-export', this.txtExport, 'export', 'insertAfter', $menu.find('#fm-btn-elx-share'));
            this.items['elx-close'] = addItem('fm-btn-elx-close', 'elx-close', this.txtClose, 'close', 'insertAfter', $menu.find('#fm-btn-elx-export'));
            this.items['elx-close'].on('click', function () { me.onClose(); });
            // About (legal notices) used to sit in the left icon rail, which the window no longer has
            this.items['elx-about'] = addItem('fm-btn-elx-about', 'elx-about', this.txtAbout, 'info', 'insertBefore', $menu.find('#fm-btn-settings'));
            this.items['elx-about'].on('click', function () { me.onAbout(); });

            // pages
            var addPage = function (action, cls) {
                var $box = $('<div id="panel-' + action + '" class="content-box elx-page ' + cls + '" style="display:none;"></div>');
                $ctx.append($box);
                var page = new Page({ el: $box, backstage: me, menu: fm });
                fm.panels[action] = page;
                return page;
            };
            this.pages = {};
            this.pages['elx-home'] = this.renderHome(addPage('elx-home', 'elx-home'));
            this.pages['elx-open'] = this.renderOpen(addPage('elx-open', 'elx-open'));
            this.pages['elx-share'] = this.renderShare(addPage('elx-share', 'elx-share'));
            this.pages['elx-export'] = this.renderExport(addPage('elx-export', 'elx-export'));

            this.afterApplyMode();
        },

        afterApplyMode: function () {
            if (!this.rendered) return;
            var fm = this.fm, mode = fm.mode || {}, me = this;
            this.mode = mode;
            var el = function (id) { return fm.$el.find(id); };

            // captions in Word's words
            fm.miNew && fm.miNew.setCaption(this.txtNew);
            fm.miInfo && fm.miInfo.setCaption(this.txtInfo);
            fm.miSaveCopyAs && fm.miSaveCopyAs.setCaption(this.txtSaveCopy);
            fm.miDownload && fm.miDownload.setCaption(mode.canDownload ? this.txtSaveAs : this.txtDownload);
            fm.miPrint && fm.miPrint.setCaption(this.txtPrint);
            fm.miPrintWithPreview && fm.miPrintWithPreview.setCaption(this.txtPrint);
            fm.miSettings && fm.miSettings.setCaption(this.txtOptions);
            fm.miHelp && fm.miHelp.setCaption(this.txtHelp);
            fm.miClose && fm.miClose.setCaption(this.txtBack);

            // pages that moved elsewhere: Open Recent -> Open, Rename -> title bar and
            // Save As, Protect and Version History -> Info tiles, Access Rights -> Share,
            // Go to Documents / Close File -> Close
            [fm.miRecent, fm.miRename, fm.miProtect, fm.miHistory, fm.miAccess, fm.miBack].forEach(function (mi) { mi && mi.hide(); });
            el('#fm-btn-close').hide();

            this.items['elx-home'][(mode.canCreateNew || mode.canOpenRecent) ? 'show' : 'hide']();
            this.items['elx-open'][(mode.canOpenRecent || this.hostConfig().uploadUrl || this.hostConfig().openUrl) ? 'show' : 'hide']();
            this.items['elx-share'][this.canShare() ? 'show' : 'hide']();
            this.items['elx-export'][(mode.canDownload && !mode.isOffline) ? 'show' : 'hide']();
            this.items['elx-close'][(mode.canCloseEditor || mode.canBack || Common.Controllers.Desktop.isActive()) ? 'show' : 'hide']();
            if (Common.Controllers.Desktop.isActive()) this.items['elx-close'].hide(); // the desktop app keeps its Exit

            // Fluent icons for every entry
            _.each(fm.items, function (item) {
                var name = me.iconFor(item.options.action);
                if (!name || !item.$el) return;
                var $a = item.$el.children('a');
                if ($a.find('.elx-icon').length) return;
                $a.find('.menu-item-icon, .menu__icon').filter(':not(.elx-icon)').remove();
                $a.prepend(me.icon(name, 20, 'menu-item-icon uni-scale'));
            });

            this.enhanceNew();
            this.enhanceSaveAs();
            this.enhanceInfo();
            this.enhancePrint();
            fm.hideDividers && fm.hideDividers();
        },

        iconFor: function (action) {
            return ({
                'back': 'back', 'elx-home': 'home', 'new': 'new', 'elx-open': 'open', 'recent': 'open',
                'info': 'info', 'save': 'save', 'save-copy': 'save-copy', 'saveas': 'save-as', 'save-desktop': 'save-as',
                'print': 'print', 'printpreview': 'print', 'elx-share': 'share', 'rights': 'people', 'elx-export': 'export',
                'elx-close': 'close', 'opts': 'options', 'help': 'help', 'edit': 'edit', 'file:open': 'open',
                'file:exit': 'close', 'close-editor': 'close', 'exit': 'folder', 'history': 'history',
                'protect': 'lock', 'rename': 'rename', 'switch:mobile': 'window-new', 'elx-about': 'info'
            })[action];
        },

        hostConfig: function () {
            var main = ElixirShared.app.getController('Main');
            return (main && main.editorConfig && main.editorConfig.elixir) || {};
        },

        canShare: function () {
            var mode = this.mode, h = this.hostConfig();
            return !!(mode && !mode.isOffline && (mode.canRequestSharingSettings || mode.sharingSettingsUrl || h.shareUrl || this.fm.panels && this.fm.panels['rights'] && this.fm.miAccess && this.fm.document && this.fm.document.info && this.fm.document.info.sharingSettings && this.fm.document.info.sharingSettings.length));
        },

        select: function (action) { this.fm.selectMenu(action); },

        onAbout: function () {
            var lm = ElixirShared.app.getController('LeftMenu').leftMenu;
            this.fm.hide();
            lm && lm.btnAbout && lm.btnAbout.toggle(true);
        },

        onClose: function () {
            var mode = this.mode || {};
            this.fm.hide();
            if (mode.canCloseEditor) Common.NotificationCenter.trigger('close');
            else if (mode.canBack) Common.NotificationCenter.trigger('goback');
        },

        // ---- shared bits ------------------------------------------------------------------

        greeting: function () {
            var h = new Date().getHours();
            return h < 5 ? this.txtGoodEvening : h < 12 ? this.txtGoodMorning : h < 18 ? this.txtGoodAfternoon : this.txtGoodEvening;
        },

        formatDate: function (iso) {
            if (!iso) return '';
            var d = new Date(iso);
            if (isNaN(d.getTime())) return '';
            var lang = (Common.Locale.getCurrentLanguage && Common.Locale.getCurrentLanguage()) || 'en';
            var now = new Date(), start = new Date(now.getFullYear(), now.getMonth(), now.getDate()).getTime();
            try {
                var time = new Intl.DateTimeFormat(lang, { hour: '2-digit', minute: '2-digit' }).format(d);
                if (d.getTime() >= start) return Common.Utils.String.format(this.txtToday, time);
                if (d.getTime() >= start - 86400000) return Common.Utils.String.format(this.txtYesterday, time);
                return new Intl.DateTimeFormat(lang, { day: 'numeric', month: 'short', year: d.getFullYear() === now.getFullYear() ? undefined : 'numeric' }).format(d);
            } catch (e) { return d.toLocaleDateString(); }
        },

        fileIcon: function (title) {
            var ext = ((title || '').split('.').pop() || '').toLowerCase();
            var cls = /^(xlsx|xls|ods|csv|xlsm|xltx|ots)$/.test(ext) ? 'cell' : /^(pptx|ppt|odp|ppsx|potx|otp)$/.test(ext) ? 'slide' : /^pdf$/.test(ext) ? 'pdf' : 'word';
            return '<span class="elx-file-icon elx-file-' + cls + '">' + this.icon(cls === 'pdf' ? 'pdf' : 'document', 24) + '</span>';
        },

        recentList: function (limit, filter) {
            var recent = (this.mode && this.mode.recent) || [], me = this;
            filter = (filter || '').toLowerCase();
            var rows = recent.filter(function (r) { return !filter || (r.title || '').toLowerCase().indexOf(filter) >= 0; });
            limit && (rows = rows.slice(0, limit));
            if (!rows.length) return '<div class="elx-empty">' + esc(filter ? this.txtNoMatch : this.txtNoRecent) + '</div>';
            return '<div class="elx-filelist" role="list">' +
                '<div class="elx-filelist-head"><span class="c-name">' + esc(this.txtName) + '</span><span class="c-date">' + esc(this.txtModified) + '</span></div>' +
                rows.map(function (r) {
                    return '<div class="elx-filerow" role="listitem" tabindex="0" data-url="' + esc(r.url) + '">' +
                        me.fileIcon(r.title) +
                        '<span class="c-name"><span class="t">' + esc(r.title) + '</span><span class="f">' + esc(r.folder || '') + '</span></span>' +
                        '<span class="c-date">' + esc(me.formatDate(r.modified)) + '</span></div>';
                }).join('') + '</div>';
        },

        bindRecent: function ($root) {
            var fm = this.fm;
            $root.off('.elxrecent').on('click.elxrecent', '.elx-filerow', function () {
                fm.fireEvent('recent:open', [fm, $(this).attr('data-url')]);
            }).on('keydown.elxrecent', '.elx-filerow', function (e) {
                if (e.keyCode === Common.UI.Keys.RETURN || e.keyCode === Common.UI.Keys.SPACE) { $(this).click(); e.preventDefault(); }
            });
        },

        thumbs: function () {
            var mode = this.mode || {}, list = [];
            if (mode.canRequestCreateNew || mode.createUrl)
                list.push({ blank: true, title: this[ElixirShared.opt('blankKey', 'txtBlankDocument')] || this.txtBlankDocument });
            (mode.templates || []).forEach(function (t) {
                if (t && t.url && !/[?&]sample=true/.test(t.url) && t.title !== 'Blank') list.push(t);
            });
            return list;
        },

        thumbsHtml: function (list) {
            var me = this;
            return '<div class="elx-thumbs">' + list.map(function (t) {
                return '<button type="button" class="elx-thumb" ' + (t.blank ? 'data-blank="1"' : 'data-url="' + esc(t.url) + '"') + ' title="' + esc(t.title) + '">' +
                    '<span class="elx-thumb-img">' + (t.blank ? '<span class="elx-thumb-blank"></span>' :
                        (t.image ? '<img alt="" src="' + esc(t.image) + '">' : me.icon('template', 24))) + '</span>' +
                    '<span class="elx-thumb-title">' + esc(t.title) + '</span></button>';
            }).join('') + '</div>';
        },

        bindThumbs: function ($root) {
            var fm = this.fm;
            $root.off('.elxthumb').on('click.elxthumb', '.elx-thumb', function () {
                var $t = $(this);
                fm.fireEvent('create:new', [fm, $t.attr('data-blank') ? 'blank' : $t.attr('data-url')]);
            });
        },

        tile: function (o) {
            return '<button type="button" class="elx-tile' + (o.cls ? ' ' + o.cls : '') + '" data-action="' + o.action + '"' + (o.disabled ? ' disabled' : '') + '>' +
                '<span class="elx-tile-icon">' + this.icon(o.icon, 24) + '</span>' +
                '<span class="elx-tile-body"><span class="elx-tile-title">' + esc(o.title) + '</span>' +
                (o.desc ? '<span class="elx-tile-desc">' + esc(o.desc) + '</span>' : '') + '</span>' +
                (o.chevron ? '<span class="elx-tile-chevron">' + this.icon('chevron-down', 20) + '</span>' : '') +
                '</button>';
        },

        toast: function (text) {
            var $t = $('<div class="elx-toast" role="status"></div>').text(text).appendTo(this.fm.$el);
            setTimeout(function () { $t.addClass('out'); }, 1800);
            setTimeout(function () { $t.remove(); }, 2300);
        },

        // ---- Home -------------------------------------------------------------------------

        renderHome: function (page) {
            var me = this;
            page.refresh = function () {
                var thumbs = me.thumbs();
                page.$el.html(
                    '<div class="elx-page-inner">' +
                    '<h1 class="header elx-greeting">' + esc(me.greeting()) + '</h1>' +
                    (thumbs.length ? '<section class="elx-section"><h2>' + esc(me.txtNew) + '</h2>' + me.thumbsHtml(thumbs.slice(0, 6)) +
                        (me.isItemVisible('new') ? '<button type="button" class="elx-link" data-go="new">' + esc(me.txtMoreTemplates) + me.icon('chevron-right', 20) + '</button>' : '') +
                        '</section>' : '') +
                    (me.mode.canOpenRecent ? '<section class="elx-section"><h2>' + esc(me.txtRecent) + '</h2>' + me.recentList(12) +
                        (me.isItemVisible('elx-open') ? '<button type="button" class="elx-link" data-go="elx-open">' + esc(me.txtMoreDocuments) + me.icon('chevron-right', 20) + '</button>' : '') +
                        '</section>' : '') +
                    '</div>');
                me.bindThumbs(page.$el);
                me.bindRecent(page.$el);
                page.$el.find('[data-go]').on('click', function () { me.select($(this).attr('data-go')); });
                page.setScroller();
                page.scroller.update();
            };
            return page;
        },

        // ---- Open -------------------------------------------------------------------------

        renderOpen: function (page) {
            var me = this;
            page.refresh = function () {
                var h = me.hostConfig();
                page.$el.html(
                    '<div class="elx-page-inner">' +
                    '<h1 class="header">' + esc(me.txtOpen) + '</h1>' +
                    '<div class="elx-open-cols">' +
                        '<div class="elx-open-places" role="tablist">' +
                            '<button type="button" class="elx-place active" role="tab" aria-selected="true">' + me.icon('clock', 24) + '<span>' + esc(me.txtRecent) + '</span></button>' +
                            (h.uploadUrl ? '<button type="button" class="elx-place" data-act="browse">' + me.icon('upload', 24) + '<span>' + esc(me.txtBrowse) + '<small>' + esc(me.txtBrowseDesc) + '</small></span></button>' : '') +
                            (h.openUrl || h.homeUrl ? '<button type="button" class="elx-place" data-act="start">' + me.icon('folder', 24) + '<span>' + esc(h.folder || me.txtMyFiles) + '<small>' + esc(me.txtStartPageDesc) + '</small></span></button>' : '') +
                        '</div>' +
                        '<div class="elx-open-list">' +
                            '<div class="elx-search-row"><span class="elx-search-ic">' + me.icon('search', 20) + '</span><input type="search" class="form-control elx-open-filter" placeholder="' + esc(me.txtSearchRecent) + '" aria-label="' + esc(me.txtSearchRecent) + '"></div>' +
                            '<div class="elx-open-files">' + me.recentList(0) + '</div>' +
                        '</div>' +
                    '</div>' +
                    '<input type="file" class="elx-file-input" style="display:none" accept="' + esc((h.uploadExts || []).map(function (e) { return '.' + e; }).join(',')) + '">' +
                    '</div>');
                me.bindRecent(page.$el);
                page.$el.find('.elx-open-filter').on('input', function () {
                    page.$el.find('.elx-open-files').html(me.recentList(0, this.value));
                }).on('keydown', function (e) { e.stopPropagation(); });
                page.$el.find('[data-act=browse]').on('click', function () { page.$el.find('.elx-file-input').click(); });
                page.$el.find('[data-act=start]').on('click', function () {
                    var w = window.open(h.openUrl || h.homeUrl, '_blank');
                    w && w.focus();
                });
                page.$el.find('.elx-file-input').on('change', function () {
                    this.files && this.files[0] && me.upload(this.files[0]);
                    this.value = '';
                });
                page.setScroller();
                page.scroller.update();
            };
            return page;
        },

        upload: function (file) {
            var h = this.hostConfig(), me = this;
            var ext = (file.name.split('.').pop() || '').toLowerCase();
            if (!h.uploadUrl || h.uploadExts && h.uploadExts.indexOf(ext) < 0) {
                // formats that need a conversion are opened by the start page
                var w = window.open(h.openUrl || h.homeUrl, '_blank');
                w && w.focus();
                return;
            }
            // open the tab now (popup blockers), point it at the document once uploaded
            var win = window.open('', '_blank');
            var fd = new FormData();
            fd.append('uploadedFile', file, file.name);
            me.toast(Common.Utils.String.format(me.txtUploading, file.name));
            fetch(h.uploadUrl, { method: 'POST', body: fd, credentials: 'include' }).then(function (r) { return r.json(); }).then(function (res) {
                if (!res || res.error || !res.filename) throw new Error(res && res.error || 'upload');
                var url = (h.editorUrl || '').replace('{name}', encodeURIComponent(res.filename));
                if (win) win.location.href = url; else window.open(url, '_blank');
            }).catch(function () {
                win && win.close();
                Common.UI.warning({ msg: me.txtUploadError });
            });
        },

        // ---- Share ------------------------------------------------------------------------

        renderShare: function (page) {
            var me = this;
            page.refresh = function () {
                var mode = me.mode, h = me.hostConfig(),
                    canInvite = mode.canRequestSharingSettings || mode.sharingSettingsUrl,
                    rights = me.fm.document && me.fm.document.info && me.fm.document.info.sharingSettings;
                page.$el.html(
                    '<div class="elx-page-inner">' +
                    '<h1 class="header">' + esc(me.txtShare) + '</h1>' +
                    '<div class="elx-tiles">' +
                        (canInvite ? me.tile({ action: 'invite', icon: 'people-add', title: me.txtShareWithPeople, desc: me.txtShareWithPeopleDesc }) : '') +
                        (h.shareUrl ? me.tile({ action: 'copylink', icon: 'link', title: me.txtCopyLink, desc: me.txtCopyLinkDesc }) : '') +
                        (h.shareUrl ? me.tile({ action: 'email', icon: 'mail', title: me.txtEmailLink, desc: me.txtEmailLinkDesc }) : '') +
                        (mode.canDownload ? me.tile({ action: 'download', icon: 'download', title: me.txtSendCopy, desc: me.txtSendCopyDesc }) : '') +
                    '</div>' +
                    (rights && rights.length ? '<section class="elx-section"><h2>' + esc(me.txtPeopleWithAccess) + '</h2><div class="elx-people">' +
                        rights.map(function (r) {
                            return '<div class="elx-person">' + me.icon('person', 20) + '<span class="n">' + esc(r.user || '') + '</span><span class="p">' + esc(r.permissions || '') + '</span></div>';
                        }).join('') + '</div></section>' : '') +
                    '</div>');
                page.$el.find('[data-action=invite]').on('click', function () {
                    if (mode.canRequestSharingSettings) Common.Gateway.requestSharingSettings();
                    else if (mode.sharingSettingsUrl) Common.NotificationCenter.trigger('collaboration:sharing');
                    me.fm.hide();
                });
                page.$el.find('[data-action=copylink]').on('click', function () { me.copyLink(); });
                page.$el.find('[data-action=email]').on('click', function () {
                    var title = me.fm.document && me.fm.document.title || '';
                    window.open('mailto:?subject=' + encodeURIComponent(title) + '&body=' + encodeURIComponent(h.shareUrl), '_self');
                });
                page.$el.find('[data-action=download]').on('click', function () {
                    me.fm.fireEvent('saveas:format', [me.fm, ElixirShared.opt('defaultFormat', Asc.c_oAscFileType.DOCX)]);
                });
                page.setScroller();
                page.scroller.update();
            };
            return page;
        },

        copyLink: function () {
            var url = this.hostConfig().shareUrl, me = this;
            if (!url) return;
            var done = function () { me.toast(me.txtLinkCopied); };
            if (navigator.clipboard && navigator.clipboard.writeText) {
                navigator.clipboard.writeText(url).then(done, function () { me.copyFallback(url); done(); });
            } else { this.copyFallback(url); done(); }
        },

        copyFallback: function (text) {
            var ta = document.createElement('textarea');
            ta.value = text; ta.style.position = 'fixed'; ta.style.opacity = '0';
            document.body.appendChild(ta); ta.select();
            try { document.execCommand('copy'); } catch (e) {}
            document.body.removeChild(ta);
        },

        // ---- Export -----------------------------------------------------------------------

        formatsFor: function (groups) {
            var ft = ((this.fm.document && this.fm.document.fileType) || ElixirShared.opt('defaultExt', 'docx')).toLowerCase();
            return FORMATS().filter(function (f) {
                return groups.indexOf(f.group) >= 0 && (!f.onlyFor || f.onlyFor === ft);
            });
        },

        formatCards: function (list) {
            var me = this;
            return '<div class="elx-formats">' + list.map(function (f) {
                return '<button type="button" class="elx-format btn-doc-format" format="' + f.type + '" format-ext="' + f.ext + '">' +
                    '<span class="elx-format-icon">' + me.icon(f.icon, 24) + '</span>' +
                    '<span class="elx-format-body"><span class="elx-format-name">' + esc(me.L(f.name)) + '</span>' +
                    '<span class="elx-format-desc">' + esc(me.L(f.desc)) + '</span></span></button>';
            }).join('') + '</div>';
        },

        renderExport: function (page) {
            var me = this;
            page.refresh = function (opts) {
                var tab = page.tab || 'pdf';
                page.$el.html(
                    '<div class="elx-page-inner">' +
                    '<h1 class="header">' + esc(me.txtExport) + '</h1>' +
                    '<div class="elx-open-cols">' +
                        '<div class="elx-open-places" role="tablist">' +
                            '<button type="button" class="elx-place' + (tab === 'pdf' ? ' active' : '') + '" data-tab="pdf" role="tab">' + me.icon('pdf', 24) + '<span>' + esc(me.txtCreatePdf) + '</span></button>' +
                            '<button type="button" class="elx-place' + (tab === 'type' ? ' active' : '') + '" data-tab="type" role="tab">' + me.icon('export', 24) + '<span>' + esc(me.txtChangeType) + '</span></button>' +
                        '</div>' +
                        '<div class="elx-export-detail">' + (tab === 'pdf' ?
                            '<h2>' + esc(me.txtCreatePdf) + '</h2>' +
                            '<ul class="elx-bullets"><li>' + esc(me[ElixirShared.opt('pdfBullet1Key', 'txtPdfBullet1')]) + '</li><li>' + esc(me.txtPdfBullet2) + '</li><li>' + esc(me.txtPdfBullet3) + '</li></ul>' +
                            '<div class="elx-export-actions">' +
                                '<button type="button" class="btn normal dlg-btn primary elx-big-btn" data-fmt="pdf">' + me.icon('pdf', 24) + '<span>' + esc(me.txtCreatePdfBtn) + '</span></button>' +
                                '<button type="button" class="btn normal dlg-btn elx-big-btn" data-fmt="pdfa">' + '<span>' + esc(me.txtCreatePdfaBtn) + '</span></button>' +
                            '</div>'
                            :
                            '<h2>' + esc(me[ElixirShared.opt('docTypesKey', 'txtDocumentTypes')]) + '</h2>' + me.formatCards(me.formatsFor(['doc'])) +
                            '<h2>' + esc(me.txtOtherTypes) + '</h2>' + me.formatCards(me.formatsFor(['other', 'pdf'])) +
                            '<h2>' + esc(me.txtImageTypes) + '</h2>' + me.formatCards(me.formatsFor(['image']))
                        ) + '</div>' +
                    '</div></div>');
                page.$el.find('[data-tab]').on('click', function () { page.tab = $(this).attr('data-tab'); page.refresh(); });
                page.$el.find('[data-fmt]').on('click', function () {
                    me.fm.fireEvent('saveas:format', [me.fm, $(this).attr('data-fmt') === 'pdfa' ? Asc.c_oAscFileType.PDFA : Asc.c_oAscFileType.PDF]);
                });
                page.$el.find('.elx-format').on('click', function () {
                    me.fm.fireEvent('saveas:format', [me.fm, parseInt($(this).attr('format'))]);
                });
                page.setScroller();
                page.scroller.update();
            };
            return page;
        },

        // ---- New: blank document and the host's templates, large thumbnails ---------------

        enhanceNew: function () {
            var panel = this.fm.panels && this.fm.panels['new'];
            if (!panel || !panel.$el || panel.$el.hasClass('elx-new')) return;
            panel.$el.addClass('elx-new elx-page');
            panel.$el.html('<div class="elx-page-inner"><h1 class="header">' + esc(this.txtNew) + '</h1>' + this.thumbsHtml(this.thumbs()) + '</div>');
            this.bindThumbs(panel.$el);
            panel.scroller && panel.scroller.update();
        },

        // ---- Save As: the download formats as a Word-like list ----------------------------

        enhanceSaveAs: function () {
            var panel = this.fm.panels && this.fm.panels['saveas'], me = this;
            if (!panel || !panel.$el || panel.$el.hasClass('elx-saveas')) return;
            panel.$el.addClass('elx-saveas elx-page');
            var mode = this.mode || {};
            panel.$el.html('<div class="elx-page-inner"><h1 class="header">' + esc(this.txtSaveAs) + '</h1>' +
                '<p class="elx-lead">' + esc(this.txtSaveAsLead) + '</p>' +
                ((mode.canRename && !mode.isDesktopApp) ? '<div class="elx-tiles">' + this.tile({ action: 'rename', icon: 'rename', title: this.txtRename, desc: this.txtRenameDesc }) + '</div>' : '') +
                '<h2>' + esc(this.txtDownloadCopy) + '</h2>' +
                this.formatCards(this.formatsFor(['doc', 'pdf', 'other', 'image'])) + '</div>');
            panel.$el.find('.elx-format').on('click', function () {
                me.fm.fireEvent('saveas:format', [me.fm, parseInt($(this).attr('format'))]);
            });
            panel.$el.find('[data-action=rename]').on('click', function () {
                me.fm.hide();
                ElixirShared.app.getController('LeftMenu').clickMenuFileItem(null, 'rename');
            });
            panel.scroller && panel.scroller.update();
        },

        // ---- Info: action tiles on the left, properties column on the right --------------

        enhanceInfo: function () {
            var panel = this.fm.panels && this.fm.panels['info'], me = this;
            if (!panel || !panel.$el || !panel.$el.find('table.main').length) return;
            if (!panel.$el.hasClass('elx-info')) {
                panel.$el.addClass('elx-info elx-page');
                var $table = panel.$el.find('table.main').first();
                var $wrap = $('<div class="elx-page-inner elx-info-wrap">' +
                    '<h1 class="header">' + esc(this.txtInfo) + '</h1>' +
                    '<div class="elx-info-file"></div>' +
                    '<div class="elx-info-cols"><div class="elx-info-actions"></div><aside class="elx-info-props"><h2>' + esc(this.txtProperties) + '</h2></aside></div>' +
                    '</div>');
                $table.before($wrap);
                $wrap.find('.elx-info-props').append($table);
                panel.$el.find('.elx-info-props table.main > tbody:first-child').hide(); // the old "Document Info" title
                var $headerCell = $table.find('td.header');
                $headerCell.closest('tbody').hide();
            }
            this.refreshInfoTiles(panel);
            if (!panel._elxShow) {
                panel._elxShow = true;
                var show = panel.show;
                panel.show = function () {
                    var r = show.apply(this, arguments);
                    me.refreshInfoTiles(this);
                    return r;
                };
            }
        },

        refreshInfoTiles: function (panel) {
            var me = this, fm = this.fm, mode = this.mode || {}, doc = fm.document || {};
            var h = this.hostConfig();
            panel.$el.find('.elx-info-file').html(
                '<span class="elx-info-name">' + esc(doc.title || '') + '</span>' +
                '<span class="elx-info-loc">' + esc(h.folder || '') + '</span>');
            var protCtrl = ElixirShared.opt('protectController', 'DocProtection'),
                canProtect = !!(fm.panels['protect'] || protCtrl && ElixirShared.app.getController(protCtrl) && mode.isEdit && mode.canProtect !== false),
                canHistory = !!(mode.canUseHistory && !mode.isDisconnected);
            panel.$el.find('.elx-info-actions').html(
                this.tile({ action: 'protect', icon: 'protect', title: this.txtProtectDocument, desc: this.txtProtectDocumentDesc, cls: 'elx-tile-big', chevron: true, disabled: !canProtect }) +
                this.tile({ action: 'inspect', icon: 'inspect', title: this.txtInspectDocument, desc: this.txtInspectDocumentDesc, cls: 'elx-tile-big' }) +
                this.tile({ action: 'history', icon: 'history', title: this.txtVersionHistory, desc: canHistory ? this.txtVersionHistoryDesc : this.txtVersionHistoryNone, cls: 'elx-tile-big', disabled: !canHistory }));
            var $actions = panel.$el.find('.elx-info-actions');
            $actions.find('[data-action=protect]').on('click', function (e) { me.showProtectMenu($(this), e); });
            $actions.find('[data-action=inspect]').on('click', function () { me.inspect(); });
            $actions.find('[data-action=history]').on('click', function () {
                fm.hide();
                ElixirShared.app.getController('LeftMenu').clickMenuFileItem(null, 'history');
            });
        },

        showProtectMenu: function ($btn, e) {
            var me = this, fm = this.fm, mode = this.mode || {};
            if (!this.protectMenu) {
                var items = [];
                if (fm.panels['protect'] && mode.isPasswordSupport)
                    items.push({ caption: this.txtEncryptPassword, value: 'password' });
                if (mode.isEdit && ElixirShared.opt('protectController', 'DocProtection') && ElixirShared.app.getController(ElixirShared.opt('protectController', 'DocProtection')))
                    items.push({ caption: this.txtRestrictEditing, value: 'restrict' });
                if (fm.panels['protect'] && mode.isSignatureSupport)
                    items.push({ caption: this.txtAddSignature, value: 'signature' });
                if (!items.length) return;
                this.protectMenu = new Common.UI.Menu({ cls: 'shifted-right', menuAlign: 'tl-bl', items: items });
                this.protectMenu.on('item:click', function (menu, item) {
                    if (item.value === 'password' || item.value === 'signature') me.select('protect');
                    else {
                        fm.hide();
                        var prot = ElixirShared.app.getController(ElixirShared.opt('protectController', 'DocProtection'));
                        prot && prot.onProtectDocClick && prot.onProtectDocClick(true);
                    }
                });
            }
            var $holder = this.protectHolder || (this.protectHolder = $('<div class="elx-menu-holder" style="position:absolute;"></div>').appendTo(fm.$el));
            var off = $btn.offset(), base = fm.$el.offset();
            $holder.css({ left: off.left - base.left, top: off.top - base.top + $btn.outerHeight(), width: 1, height: 1 });
            if (!this.protectMenu.rendered) this.protectMenu.render($holder);
            this.protectMenu.show();
            e && e.stopPropagation();
        },

        // Inspect Document: what a reviewer would find (comments, tracked changes, author)
        inspect: function () {
            var api = ElixirShared.app.getController('Main').api, me = this;
            var comments = this.controller.getApplication().getCollection('Common.Collections.Comments');
            var nComments = comments ? comments.filter(function (c) { return !c.get('hide'); }).length : 0;
            var nChanges = 0;
            try {
                var rc = ElixirShared.app.getController('Common.Controllers.ReviewChanges');
                nChanges = api.asc_GetRevisionsChangesStack ? (api.asc_GetRevisionsChangesStack() || []).length : 0;
                if (!nChanges && rc && rc._state && rc._state.changesCount) nChanges = rc._state.changesCount;
            } catch (e) {}
            var props = api.asc_getCoreProps ? api.asc_getCoreProps() : null;
            var author = props ? (props.asc_getCreator && props.asc_getCreator()) || '' : '';
            var rows = [
                { icon: 'comments', title: this.txtInspComments, desc: nComments ? Common.Utils.String.format(this.txtInspCommentsFound, nComments) : this.txtInspNone, act: nComments && this.mode.canDeleteComments !== false ? 'comments' : '' },
                { icon: 'review', title: this.txtInspChanges, desc: nChanges ? Common.Utils.String.format(this.txtInspChangesFound, nChanges) : this.txtInspNone, act: '' },
                { icon: 'person', title: this.txtInspProps, desc: author ? Common.Utils.String.format(this.txtInspAuthor, author) : this.txtInspNone, act: '' }
            ];
            var html = '<div class="elx-inspect">' + rows.map(function (r) {
                return '<div class="elx-inspect-row">' + me.icon(r.icon, 24) + '<div class="b"><div class="t">' + esc(r.title) + '</div><div class="d">' + esc(r.desc) + '</div></div>' +
                    (r.act ? '<button type="button" class="btn normal dlg-btn" data-remove="' + r.act + '">' + esc(me.txtRemoveAll) + '</button>' : '') + '</div>';
            }).join('') + '</div>';
            var win = new Common.UI.Window({
                title: this.txtInspectDocument,
                width: 520,
                cls: 'modal-dlg elx-inspect-dlg',
                tpl: '<div class="box" style="padding: 8px 16px 16px;">' + html + '</div><div class="footer center"><button class="btn normal dlg-btn primary" result="ok">' + esc(this.txtClose) + '</button></div>',
                buttons: null
            });
            win.show();
            win.$window.find('[result=ok]').on('click', function () { win.close(); });
            win.$window.find('[data-remove=comments]').on('click', function () {
                api.asc_RemoveAllComments && api.asc_RemoveAllComments(false, false);
                $(this).prop('disabled', true).text(me.txtRemoved);
            });
        },

        // ---- Print: the preview page in the browser, Word-like settings column ------------

        enhancePrint: function () {
            var panel = this.fm.panels && this.fm.panels['printpreview'];
            if (!panel || !panel.$el || !panel.cmbPrinter || panel._elx) return;
            panel._elx = true;
            var mode = this.mode || {}, me = this;
            panel.$el.addClass('elx-print');
            if (!mode.isDesktopApp) {
                // the browser picks the printer, the copies and the sides in its own dialog
                var hideRow = function (sel) {
                    var $c = panel.$el.find(sel).closest('tr');
                    $c.hide(); $c.prev('tr').hide();
                };
                hideRow('#print-combo-printer');
                hideRow('#print-combo-color-printing');
                hideRow('#print-combo-sides');
                panel.$el.find('#print-txt-copies').closest('tr').hide();
                panel.$el.find('#print-btn-system-dialog').closest('tr').hide();
                panel.btnPrint.setDisabled(false);
                var $printerRow = $('<tr class="elx-print-printer"><td class="padding-large"><div class="elx-printer-card">' + this.icon('print', 24) +
                    '<span><b>' + esc(this.txtBrowserPrinter) + '</b><small>' + esc(this.txtBrowserPrinterDesc) + '</small></span></div></td></tr>');
                panel.$el.find('#print-combo-range').closest('tr').prev('tr').before($printerRow);
            }
            // Word puts the Print button at the top of the settings column
            var $btns = panel.$el.find('.fms-btn-apply');
            $btns.addClass('elx-print-actions');
            panel.$el.find('.main-header').after($('<div class="elx-print-top"></div>').append($btns.find('.footer')));
            $btns.hide();
            panel.btnPrint.$el.prepend(this.icon('print', 20));
            panel.updateScroller && panel.updateScroller();
        }


    }, STRINGS, Common.Views.ElixirBackstage || {}));
});

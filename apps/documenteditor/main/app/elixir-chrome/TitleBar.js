/*
 * ElixirOffice window chrome: Word-like title bar.
 *
 * - file name with the save state ("Saved" / "Saving…") and a flyout to rename
 *   the file (when the host supports it), see its location and version history;
 * - centred "Search (Alt+Q)" box that runs ribbon commands (CommandIndex) and
 *   falls back to Find in the document;
 * - on the tab row: Comments, the Editing / Reviewing / Viewing mode, Share.
 *
 * Built on the upstream header (Common.Views.Header) instance, which keeps its
 * own buttons and events.
 *
 * Licensed under the GNU AGPL v3 like the rest of web-apps.
 */
define([
    'core',
    'documenteditor/main/app/elixir-chrome/Icons',
    'documenteditor/main/app/elixir-chrome/CommandIndex'
], function () {
    'use strict';

    var Icons = window.ElixirChromeIcons, Index = window.ElixirCommandIndex;
    var esc = function (s) { return Common.Utils.String.htmlEncode(s === undefined || s === null ? '' : String(s)); };
    var isMac = Common.Utils.isMac;

    var STRINGS = {
        txtSaved: 'Saved',
        txtSaving: 'Saving…',
        txtNotSaved: 'Not saved',
        txtOffline: 'Offline',
        txtSavedTip: 'All changes are saved',
        txtFileName: 'File name',
        txtLocation: 'Location',
        txtOpenLocation: 'Open file location',
        txtVersionHistory: 'Version history',
        txtRenameHint: 'Press Enter to rename the file',
        txtInvalidName: 'A file name cannot contain any of these characters:',
        txtSearch: 'Search',
        txtSearchPlaceholder: 'Search ({0})',
        txtSearchLabel: 'Search for commands and in the document',
        txtActions: 'Actions',
        txtSuggested: 'Suggested actions',
        txtFindInDoc: 'Find in document',
        txtFindText: 'Find “{0}”',
        txtReplaceText: 'Replace “{0}”…',
        txtNoCommand: 'No matching command',
        txtComments: 'Comments',
        txtCommentsTip: 'Show comments',
        txtShare: 'Share',
        txtShareTip: 'Share this document'
    };

    DE.Views.ElixirTitleBar = Backbone.View.extend(_.extend({

        initialize: function (options) {
            this.controller = options.controller;
            this.state = { modified: false, saving: false, offline: false };
        },

        icon: function (name, size, cls) { return Icons ? Icons.svg(name, size, cls) : ''; },
        shortcut: function () { return isMac ? '⌃⌥Q' : 'Alt+Q'; },

        setup: function (mode) {
            var main = DE.getController('Main');
            this.mode = mode || main.appOptions || {};
            this.api = main.api;
            this.header = DE.getController('Viewport').header;
            this.toolbar = DE.getController('Toolbar').toolbar;
            this.$title = $('#box-document-title');
            if (!this.$title.length || !this.header || this.ready) return;
            this.ready = true;
            $('body').addClass('elx-chrome');
            this.$title.addClass('elx-titlebar');

            this.renderDocName();
            this.renderSearch();
            this.renderRight();
            this.bindSaveState();
            this.bindShortcut();
            var me = this;
            Common.NotificationCenter.on('window:resize', function () { me.fitName(); });
            this.fitName();
        },

        hostConfig: function () {
            var main = DE.getController('Main');
            return (main && main.editorConfig && main.editorConfig.elixir) || {};
        },

        // ---- file name, save state, flyout ------------------------------------------------

        renderDocName: function () {
            var me = this, header = this.header;
            var $box = $('#id-box-doc-name');
            if (!$box.length) return;
            $box.addClass('elx-docname-box');
            this.$doc = $('<button type="button" class="elx-docname" aria-haspopup="dialog" aria-expanded="false">' +
                '<span class="elx-docname-text"></span>' +
                '<span class="elx-docstate" aria-live="polite"><span class="elx-docstate-ic"></span><span class="elx-docstate-txt"></span></span>' +
                '<span class="elx-docname-chev">' + this.icon('chevron-down', 20) + '</span></button>');
            $box.append(this.$doc);
            this.$doc.on('click', function (e) { me.toggleFlyout(); e.stopPropagation(); });

            // follow the upstream caption updates (rename, meta, modified flag)
            var setCaption = header.setDocumentCaption, setChanged = header.setDocumentChanged;
            header.setDocumentCaption = function () { var r = setCaption.apply(this, arguments); me.updateName(); return r; };
            header.setDocumentChanged = function () { var r = setChanged.apply(this, arguments); me.updateName(); return r; };
            this.updateName();
            this.updateState();
        },

        caption: function () {
            return (this.header && this.header.getDocumentCaption && this.header.getDocumentCaption()) || '';
        },

        splitName: function (name) {
            var i = name.lastIndexOf('.');
            return i > 0 ? { base: name.substring(0, i), ext: name.substring(i) } : { base: name, ext: '' };
        },

        updateName: function () {
            if (!this.$doc) return;
            var name = this.caption();
            this.$doc.find('.elx-docname-text').text(this.splitName(name).base);
            this.$doc.attr('title', name);
            this.fitName();
        },

        // the file name and its state stop before the centred search box
        fitName: function () {
            if (!this.$doc || !this.$search) return;
            var rtl = Common.UI.isRTL(), d = this.$doc[0].getBoundingClientRect(), s = this.$search[0].getBoundingClientRect();
            var room = rtl ? d.right - s.right - 12 : s.left - d.left - 12;
            // narrow windows keep the name and show the state as an icon only
            this.$doc.toggleClass('elx-compact', room < 240);
            this.$doc.css('max-width', Math.max(80, Math.floor(room)) + 'px');
        },

        bindSaveState: function () {
            var me = this, api = this.api;
            if (!api) return;
            var A = Asc.c_oAscAsyncAction;
            // "unsaved changes" = what the editor still has to send (the Save button state)
            var refresh = function () {
                me.state.modified = api.asc_isDocumentCanSave ? !!api.asc_isDocumentCanSave() : !!(api.isDocumentModified && api.isDocumentModified());
                me.updateState();
            };
            api.asc_registerCallback('asc_onDocumentModifiedChanged', refresh);
            api.asc_registerCallback('asc_onDocumentCanSaveChanged', refresh);
            var isSave = function (id) { return id === A.Save || id === A.ForceSaveButton || id === A.ForceSaveTimeout; };
            api.asc_registerCallback('asc_onStartAction', function (type, id) {
                if (isSave(id)) { me.state.saving = true; me.updateState(); }
            });
            api.asc_registerCallback('asc_onEndAction', function (type, id) {
                if (isSave(id)) { me.state.saving = false; refresh(); }
            });
            api.asc_registerCallback('asc_onCoAuthoringDisconnect', function () { me.state.offline = true; me.updateState(); });
            Common.NotificationCenter.on('api:disconnect', function () { me.state.offline = true; me.updateState(); });
        },

        updateState: function () {
            if (!this.$doc) return;
            var s = this.state, mode = this.mode || {}, me = this;
            var $st = this.$doc.find('.elx-docstate');
            if (!mode.isEdit) { $st.hide(); return; }
            var autosave = Common.Utils.InternalSettings.get('de-settings-autosave');
            autosave = autosave === undefined || autosave === null ? true : !!autosave;
            var key = s.offline ? 'offline' : (s.saving || s.modified && autosave) ? 'saving' : s.modified ? 'unsaved' : 'saved';
            clearTimeout(this._stateTimer);
            var apply = function () {
                var txt = { saved: me.txtSaved, saving: me.txtSaving, unsaved: me.txtNotSaved, offline: me.txtOffline }[key];
                var ic = { saved: 'saved', saving: 'saving', unsaved: 'unsaved', offline: 'unsaved' }[key];
                $st.show().attr('data-state', key).attr('title', key === 'saved' ? me.txtSavedTip : txt);
                $st.find('.elx-docstate-ic').html(me.icon(ic, 20));
                $st.find('.elx-docstate-txt').text(txt);
            };
            // "Saved" settles a moment after the last change, like Word
            if (key === 'saved' && $st.attr('data-state') === 'saving') this._stateTimer = setTimeout(apply, 700);
            else apply();
        },

        toggleFlyout: function () {
            this.$flyout && this.$flyout.is(':visible') ? this.closeFlyout() : this.openFlyout();
        },

        openFlyout: function () {
            var me = this, mode = this.mode || {}, h = this.hostConfig();
            var name = this.caption(), parts = this.splitName(name);
            var canRename = !!(mode.canRename && !mode.isDesktopApp && !this.header.readOnly);
            Common.UI.Menu.Manager.hideAll();
            this.closeSearch();
            if (!this.$flyout) {
                this.$flyout = $('<div class="elx-flyout elx-docflyout" role="dialog"></div>').appendTo(document.body);
                $(document).on('mousedown.elxfly', function (e) {
                    if (me.$flyout && me.$flyout.is(':visible') && !$(e.target).closest('.elx-docflyout, .elx-docname').length) me.closeFlyout();
                });
            }
            var html =
                '<label class="elx-fly-label" for="elx-fname">' + esc(this.txtFileName) + '</label>' +
                '<div class="elx-fname-row"><input id="elx-fname" type="text" class="form-control" spellcheck="false" autocomplete="off" value="' + esc(parts.base) + '"' + (canRename ? '' : ' disabled') + '>' +
                    '<span class="elx-fname-ext">' + esc(parts.ext) + '</span></div>' +
                (canRename ? '<div class="elx-fly-hint">' + esc(this.txtRenameHint) + '</div>' : '') +
                (h.folder ? '<div class="elx-fly-label">' + esc(this.txtLocation) + '</div><div class="elx-fly-loc">' + this.icon('folder', 20) + '<span>' + esc(h.folder) + '</span></div>' : '') +
                '<div class="elx-fly-actions">' +
                    (mode.canBack ? '<button type="button" class="elx-fly-link" data-act="location">' + this.icon('open', 20) + '<span>' + esc(this.txtOpenLocation) + '</span></button>' : '') +
                    (mode.canUseHistory && !mode.isDisconnected ? '<button type="button" class="elx-fly-link" data-act="history">' + this.icon('history', 20) + '<span>' + esc(this.txtVersionHistory) + '</span></button>' : '') +
                '</div>';
            this.$flyout.html(html);
            var off = this.$doc.offset();
            var left = Common.UI.isRTL() ? Math.max(8, off.left + this.$doc.outerWidth() - 320) : Math.max(8, off.left);
            this.$flyout.css({ top: off.top + this.$doc.outerHeight() + 6, left: left }).show();
            this.$doc.attr('aria-expanded', 'true');
            this.api && this.api.asc_enableKeyEvents(false);

            var $in = this.$flyout.find('#elx-fname');
            var commit = function () {
                var v = $.trim($in.val());
                if (!canRename || !v || v === parts.base) return true;
                if (/[\t*+:"<>?|\\/]/.test(v)) {
                    me.$flyout.find('.elx-fly-hint').addClass('error').text(me.txtInvalidName + ' *+:"<>?|\\/');
                    return false;
                }
                me.header.fireEvent('rename', [v]);
                me.header.setDocumentCaption(v + parts.ext);
                return true;
            };
            $in.on('keydown', function (e) {
                e.stopPropagation();
                if (e.keyCode === Common.UI.Keys.RETURN) { commit() && me.closeFlyout(true); }
                else if (e.keyCode === Common.UI.Keys.ESC) { me.closeFlyout(true); }
            });
            canRename && setTimeout(function () { $in.focus().select(); }, 30);
            this.$flyout.find('[data-act=location]').on('click', function () { me.closeFlyout(true); Common.NotificationCenter.trigger('goback'); });
            this.$flyout.find('[data-act=history]').on('click', function () {
                me.closeFlyout(true);
                DE.getController('LeftMenu').clickMenuFileItem(null, 'history');
            });
            this.$flyout.on('keydown', function (e) { if (e.keyCode === Common.UI.Keys.ESC) me.closeFlyout(true); });
        },

        closeFlyout: function (focusDoc) {
            if (!this.$flyout) return;
            this.$flyout.hide();
            this.$doc && this.$doc.attr('aria-expanded', 'false');
            this.api && this.api.asc_enableKeyEvents(true);
            focusDoc && Common.NotificationCenter.trigger('edit:complete');
        },

        // ---- search box -------------------------------------------------------------------

        renderSearch: function () {
            var me = this;
            var ph = Common.Utils.String.format(this.txtSearchPlaceholder, this.shortcut());
            this.$search = $('<div class="elx-search" role="search">' +
                '<span class="elx-search-ic">' + this.icon('search', 20) + '</span>' +
                '<input type="text" class="elx-search-input" spellcheck="false" autocomplete="off" role="combobox" aria-expanded="false" aria-autocomplete="list" aria-controls="elx-search-pop" placeholder="' + esc(ph) + '" aria-label="' + esc(this.txtSearchLabel) + '">' +
                '<div id="elx-search-pop" class="elx-search-pop" role="listbox" style="display:none;"></div></div>');
            this.$title.append(this.$search);
            this.$input = this.$search.find('input');
            this.$pop = this.$search.find('.elx-search-pop');

            this.$input.on('focus', function () { me.openSearch(); });
            this.$input.on('input', function () { me.renderResults(); });
            this.$input.on('keydown', function (e) { me.onSearchKey(e); });
            this.$input.on('blur', function () { setTimeout(function () { if (!me.$search[0].contains(document.activeElement)) me.closeSearch(); }, 120); });
            this.$pop.on('mousedown', function (e) { e.preventDefault(); });
            this.$pop.on('click', '.elx-sr', function () { me.runResult($(this).index('.elx-search-pop .elx-sr')); });
            this.$pop.on('mousemove', '.elx-sr', function () { me.highlight($(this).index('.elx-search-pop .elx-sr')); });
        },

        bindShortcut: function () {
            var me = this;
            // Word's Alt+Q (Ctrl+Option+Q on macOS, the editor's convention for Alt shortcuts);
            // captured before the editor's own handlers
            window.addEventListener('keydown', function (e) {
                var q = e.keyCode === 81 || e.code === 'KeyQ';
                if (!q) return;
                var hit = isMac ? (e.ctrlKey && e.altKey && !e.metaKey && !e.shiftKey) : (e.altKey && !e.ctrlKey && !e.metaKey && !e.shiftKey);
                if (!hit) return;
                e.preventDefault();
                e.stopImmediatePropagation();
                me.focusSearch();
            }, true);
        },

        focusSearch: function () {
            if (!this.$input) return;
            var fm = DE.getController('LeftMenu').leftMenu.menuFile;
            fm && fm.isVisible() && fm.hide();
            Common.UI.Menu.Manager.hideAll();
            this.$input.focus().select();
        },

        openSearch: function () {
            this.api && this.api.asc_enableKeyEvents(false);
            Index.build(this.toolbar);
            this.$search.addClass('open');
            this.renderResults();
        },

        closeSearch: function () {
            if (!this.$search || !this.$search.hasClass('open')) return;
            this.$search.removeClass('open');
            this.$pop.hide();
            this.$input.attr('aria-expanded', 'false');
            this.api && this.api.asc_enableKeyEvents(true);
        },

        suggestions: function () {
            var tb = this.toolbar || {}, list = [];
            var want = [tb.btnInsertTable, tb.btnsPageBreak && tb.btnsPageBreak[0], tb.btnInsertShape, tb.btnInsertText, tb.btnPageMargins, tb.btnColumns];
            want.forEach(function (b) {
                if (!b) return;
                var c = _.find(Index.cmds || [], function (c) { return c.btn === b && !c.item; });
                c && list.push(c);
            });
            return list;
        },

        renderResults: function () {
            var q = $.trim(this.$input.val()), me = this, html = '';
            var cmds = q ? Index.search(q, 7) : this.suggestions();
            this.results = [];
            if (cmds.length) {
                html += '<div class="elx-sg">' + esc(q ? this.txtActions : this.txtSuggested) + '</div>';
                cmds.forEach(function (c) {
                    var dis = Index.isDisabled(c);
                    me.results.push({ type: 'cmd', c: c });
                    var where = [c.tabCaption, c.group, c.item ? c.parent : ''].filter(Boolean).join(' › ');
                    html += '<div class="elx-sr' + (dis ? ' disabled' : '') + '" role="option" id="elx-sr-' + (me.results.length - 1) + '">' +
                        '<span class="elx-sr-ic">' + me.cmdIcon(c) + '</span>' +
                        '<span class="elx-sr-label">' + esc(c.label) + '</span>' +
                        (where ? '<span class="elx-sr-where">' + esc(where) + '</span>' : '') + '</div>';
                });
            } else if (q) {
                html += '<div class="elx-sg">' + esc(this.txtActions) + '</div><div class="elx-sr-empty">' + esc(this.txtNoCommand) + '</div>';
            }
            if (q) {
                html += '<div class="elx-sg">' + esc(this.txtFindInDoc) + '</div>';
                this.results.push({ type: 'find', q: q });
                html += '<div class="elx-sr" role="option" id="elx-sr-' + (this.results.length - 1) + '"><span class="elx-sr-ic">' + this.icon('search', 20) + '</span><span class="elx-sr-label">' + esc(Common.Utils.String.format(this.txtFindText, q)) + '</span></div>';
                if (this.mode && this.mode.isEdit) {
                    this.results.push({ type: 'replace', q: q });
                    html += '<div class="elx-sr" role="option" id="elx-sr-' + (this.results.length - 1) + '"><span class="elx-sr-ic">' + this.icon('rename', 20) + '</span><span class="elx-sr-label">' + esc(Common.Utils.String.format(this.txtReplaceText, q)) + '</span></div>';
                }
            }
            this.$pop.html(html).toggle(!!html);
            this.$input.attr('aria-expanded', html ? 'true' : 'false');
            // first enabled command is the default, like Word; otherwise Find
            var first = _.findIndex(this.results, function (r) { return r.type !== 'cmd' || !Index.isDisabled(r.c); });
            this.highlight(first < 0 ? 0 : first);
        },

        cmdIcon: function (c) {
            var $ic = c.btn && c.btn.cmpEl ? c.btn.cmpEl.find('svg.icon, i.icon, .icon').first() : $();
            if ($ic.length) {
                var $cl = $ic.clone();
                $cl.removeAttr('id').addClass('elx-sr-cmdicon');
                return $('<div>').append($cl).html();
            }
            return this.icon('more', 20);
        },

        highlight: function (i) {
            var $rows = this.$pop.find('.elx-sr');
            if (!$rows.length) { this.active = -1; return; }
            i = Math.max(0, Math.min(i, $rows.length - 1));
            this.active = i;
            $rows.removeClass('active').eq(i).addClass('active');
            this.$input.attr('aria-activedescendant', 'elx-sr-' + i);
        },

        onSearchKey: function (e) {
            var K = Common.UI.Keys;
            e.stopPropagation();
            if (e.keyCode === K.DOWN) { this.highlight((this.active || 0) + 1); e.preventDefault(); }
            else if (e.keyCode === K.UP) { this.highlight((this.active || 0) - 1); e.preventDefault(); }
            else if (e.keyCode === K.RETURN) { this.runResult(this.active); e.preventDefault(); }
            else if (e.keyCode === K.ESC) {
                e.preventDefault();
                if (this.$input.val()) { this.$input.val(''); this.renderResults(); }
                else { this.closeSearch(); this.$input.blur(); Common.NotificationCenter.trigger('edit:complete'); }
            }
        },

        runResult: function (i) {
            var r = this.results && this.results[i];
            if (!r) return;
            if (r.type === 'cmd' && Index.isDisabled(r.c)) return;
            this.$input.val('');
            this.closeSearch();
            this.$input.blur();
            if (r.type === 'cmd') {
                Index.run(r.c, this.toolbar);
            } else {
                this.controller.findInDocument(r.q, r.type === 'replace');
            }
        },

        // ---- tab row: Comments, mode, Share ------------------------------------------------

        renderRight: function () {
            var me = this, mode = this.mode || {}, header = this.header;
            var $right = $('#box-right-btn-group');
            if (!$right.length) return;
            $right.addClass('elx-right');
            var mkBtn = function (cls, icon, caption, hint) {
                var $b = $('<button type="button" class="btn btn-header elx-hbtn ' + cls + '">' + me.icon(icon, 20) + '<span class="caption">' + esc(caption) + '</span></button>');
                $b.attr('data-hint', '0').attr('data-hint-direction', 'bottom').attr('data-hint-offset', 'big');
                hint && $b.tooltip({ title: hint, placement: 'cursor' });
                return $b;
            };
            if (mode.canViewComments && mode.canCoAuthoring !== false) {
                var $set = $('<div class="hedset elx-hedset-comments"></div>');
                this.$comments = mkBtn('elx-btn-comments', 'comment', this.txtComments, this.txtCommentsTip);
                $set.append(this.$comments);
                $right.find('[data-layout-name="header-editMode"]').before($set);
                this.$comments.on('click', function () { me.controller.toggleComments(); });
            }
            if (!header.btnShare && this.controller.canShare()) {
                this.$share = mkBtn('elx-btn-share', 'share', this.txtShare, this.txtShareTip);
                $right.find('#slot-btn-share').show().empty().append(this.$share);
                this.$share.on('click', function () { me.controller.openBackstage('elx-share'); });
            } else if (header.btnShare) {
                header.btnShare.cmpEl.addClass('elx-btn-share');
            }
        },

        setCommentsPressed: function (pressed) {
            this.$comments && this.$comments.toggleClass('active', !!pressed).attr('aria-pressed', !!pressed);
        }

    }, STRINGS, DE.Views.ElixirTitleBar || {}));
});

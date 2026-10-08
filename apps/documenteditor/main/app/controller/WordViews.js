/*
 * SPDX-FileCopyrightText: 2026 ElixirOffice contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 *
 * DE.Controllers.WordViews: the document views of Word (sdkjs word/Editor/word-views.js).
 *
 *   View > Read Mode (Alt, W, F)      read-only, ribbon folded, the text flows in pages of the window size;
 *                                     arrows, Page Up/Down, Space, Home, End, a page number + Enter, Esc
 *   View > Print Layout (Alt, W, P)   Ctrl+Alt+P
 *   View > Outline (Alt, W, U)        Ctrl+Alt+O; the Outlining tab (Alt, U) appears before Home
 *   View > Draft (Alt, W, E)          Ctrl+Alt+N
 * and the status hints of the field commands (Ctrl+F11, Ctrl+Shift+F11, F9 on locked fields).
 */

define([
    'core',
    'common/main/lib/component/Button',
    'common/main/lib/component/CheckBox',
    'common/main/lib/component/ComboBox'
], function () {
    'use strict';

    DE.Controllers.WordViews = Backbone.Controller.extend(_.extend({
        initialize: function () {
            var me = this;
            this._state = {view: 'print', read: false, wantRead: false, pageInput: ''};
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

        // ---- Outlining tab -----------------------------------------------------------------
        onToolbarRender: function (toolbar) {
            var mode = this.appOptions();
            if (!mode.isEdit || mode.isEditMailMerge || mode.isEditDiagram || mode.isEditOle || this.toolbar) return;
            this.toolbar = toolbar;
            var me = this,
                slot = function (id, cls) { return '<span class="btn-slot' + (cls ? ' ' + cls : '') + '" id="' + id + '"></span>'; },
                big = function (id) { return '<span class="btn-slot text x-huge" id="' + id + '"></span>'; };
            var $panel = $('<section class="panel" data-tab="outlining" role="tabpanel" aria-labelledby="outlining">' +
                '<div class="group" data-group="outline-tools">' +
                    '<div class="elset">' + slot('elx-outl-h1') + slot('elx-outl-promote') + '<span class="btn-slot" id="elx-outl-level" style="width:110px;"></span>' + slot('elx-outl-demote') + slot('elx-outl-body') + '</div>' +
                    '<div class="elset">' + slot('elx-outl-up') + slot('elx-outl-down') + slot('elx-outl-expand') + slot('elx-outl-collapse') + '</div>' +
                '</div>' +
                '<div class="separator long"></div>' +
                '<div class="group" data-group="outline-show">' +
                    '<div class="elset"><span class="btn-slot text font-size-normal margin-right-6" id="elx-outl-show-label"></span><span class="btn-slot" id="elx-outl-show" style="width:110px;"></span></div>' +
                    '<div class="elset">' + slot('elx-outl-formatting', 'text') + '</div>' +
                '</div>' +
                '<div class="separator long"></div>' +
                '<div class="group" data-group="outline-close">' + big('elx-outl-close') + '</div>' +
                '</section>');

            var lock = [Common.enumLock.lostConnect, Common.enumLock.disableOnStart, Common.enumLock.viewMode, Common.enumLock.previewReviewMode, Common.enumLock.docLockView, Common.enumLock.docLockComments];
            var btn = function (id, icon, hint, cmd) {
                var b = new Common.UI.Button({cls: 'btn-toolbar', iconCls: 'toolbar__icon ' + icon, hint: hint, lock: lock, dataHint: '1', dataHintDirection: 'bottom', dataHintOffset: 'small'});
                b.render($panel.find('#' + id));
                b.on('click', function () { me.command(cmd); });
                me.controls.push(b);
                return b;
            };
            this.controls = [];
            this.btnH1 = btn('elx-outl-h1', 'btn-firstitem', this.tipPromoteH1, 'h1');
            this.btnPromote = btn('elx-outl-promote', 'btn-decoffset', this.tipPromote, 'promote');
            this.btnDemote = btn('elx-outl-demote', 'btn-incoffset', this.tipDemote, 'demote');
            this.btnBody = btn('elx-outl-body', 'btn-lastitem', this.tipDemoteBody, 'body');
            this.btnUp = btn('elx-outl-up', 'btn-arrow-up', this.tipMoveUp, 'up');
            this.btnDown = btn('elx-outl-down', 'btn-arrow-down', this.tipMoveDown, 'down');
            this.btnExpand = btn('elx-outl-expand', 'btn-zoomup', this.tipExpand, 'expand');
            this.btnCollapse = btn('elx-outl-collapse', 'btn-zoomdown', this.tipCollapse, 'collapse');

            var levels = [];
            for (var i = 1; i <= 9; i++) levels.push({value: i - 1, displayValue: Common.Utils.String.format(this.txtLevel, i)});
            levels.push({value: 9, displayValue: this.txtBodyText});
            this.cmbLevel = new Common.UI.ComboBox({cls: 'input-group-nr', menuStyle: 'min-width: 110px;', editable: false, data: levels, lock: lock, hint: this.tipOutlineLevel, dataHint: '1', dataHintDirection: 'bottom', dataHintOffset: 'small'});
            this.cmbLevel.render($panel.find('#elx-outl-level'));
            this.cmbLevel.on('selected', function (combo, record) { me.command('setlevel', record.value); });
            this.controls.push(this.cmbLevel);

            var shows = [{value: 0, displayValue: this.txtAllLevels}];
            for (i = 1; i <= 9; i++) shows.push({value: i, displayValue: Common.Utils.String.format(this.txtLevel, i)});
            $panel.find('#elx-outl-show-label').text(this.txtShowLevel);
            this.cmbShow = new Common.UI.ComboBox({cls: 'input-group-nr', menuStyle: 'min-width: 110px;', editable: false, data: shows, lock: lock, hint: this.txtShowLevel, dataHint: '1', dataHintDirection: 'bottom', dataHintOffset: 'small'});
            this.cmbShow.render($panel.find('#elx-outl-show'));
            this.cmbShow.setValue(0);
            this.cmbShow.on('selected', function (combo, record) { me.command('level', record.value); });
            this.controls.push(this.cmbShow);

            this.chkFormatting = new Common.UI.CheckBox({labelText: this.txtShowFormatting, value: true, lock: lock, dataHint: '1', dataHintDirection: 'left', dataHintOffset: 'small'});
            this.chkFormatting.render($panel.find('#elx-outl-formatting'));
            this.chkFormatting.on('change', function (field, value) {
                var st = me.api && me.api.asc_GetOutlineState();
                if (st && !!st.showFormatting !== (value === 'checked' || value === true)) me.command('formatting');
            });
            this.controls.push(this.chkFormatting);

            this.btnClose = new Common.UI.Button({cls: 'btn-toolbar x-huge icon-top', iconCls: 'toolbar__icon btn-close', caption: this.txtCloseOutline, lock: [Common.enumLock.lostConnect, Common.enumLock.disableOnStart], dataHint: '1', dataHintDirection: 'bottom', dataHintOffset: 'small'});
            this.btnClose.render($panel.find('#elx-outl-close'));
            this.btnClose.on('click', function () { me.setView('print'); });
            this.controls.push(this.btnClose);

            toolbar.addTab({action: 'outlining', caption: this.txtOutlining, extcls: 'canedit', layoutname: 'toolbar-outlining', dataHintTitle: 'U'}, $panel, 0);
            Array.prototype.push.apply(toolbar.lockControls, this.controls);
            Common.Utils.lockControls(Common.enumLock.disableOnStart, true, {array: this.controls});
        },

        onAppReady: function () {
            var me = this, mode = this.appOptions();
            this.api = this.getApplication().getController('Viewport').getApi();
            if (!this.api || !this.api.asc_SetWordView) return;
            this.controls && Common.Utils.lockControls(Common.enumLock.disableOnStart, false, {array: this.controls});
            this._state.lastMode = mode.isEdit ? 'edit' : 'view';
            this.api.asc_registerCallback('asc_onWordViewChanged', _.bind(this.onViewChanged, this));
            this.api.asc_registerCallback('asc_onWordFieldHint', _.bind(this.onFieldHint, this));
            this.api.asc_registerCallback('asc_onFocusObject', function () { me._state.view === 'outline' && me.updateOutlineControls(); });
            Common.NotificationCenter.on('doc:mode-changed', _.bind(this.onDocModeChanged, this));
            document.addEventListener('keydown', _.bind(this.onReadKey, this), true);
            $(window).on('resize', _.debounce(function () { me._state.read && me.applyReadLayout(); }, 300));
            // the KeyTips of the new controls (WordKeyTips reads these tables when Alt is pressed)
            var kt = this.getApplication().getController('WordKeyTips');
            if (kt && kt.commands && kt.tabs) {
                kt.tabs.outlining = 'U';
                kt.commands.view = (kt.commands.view || []).concat([['#slot-elx-view-outline', 'U'], ['#slot-elx-view-draft', 'E']]);
                kt.commands.outlining = [['#elx-outl-h1', 'U'], ['#elx-outl-promote', 'P'], ['#elx-outl-level', 'L'], ['#elx-outl-demote', 'D'],
                    ['#elx-outl-body', 'B'], ['#elx-outl-up', 'M'], ['#elx-outl-down', 'O'], ['#elx-outl-expand', 'E'], ['#elx-outl-collapse', 'C'],
                    ['#elx-outl-show', 'S'], ['#elx-outl-formatting', 'T'], ['#elx-outl-close', 'C2']];
            }
            setTimeout(function () { me.arrangeViewTab(mode); }, 0);
        },

        // View tab: Outline and Draft after Read Mode and Print Layout (RibbonTabs.arrangeView)
        arrangeViewTab: function (mode) {
            var me = this, rt = this.getApplication().getController('RibbonTabs'), v = rt && rt.view,
                $group = $('#slot-elx-view-print').closest('.group');
            if (!$group.length || !v || !v.btnViewPrint || this.btnViewOutline) return;
            if (mode.isEdit) {
                $group.append('<span class="btn-slot text x-huge" id="slot-elx-view-outline"></span><span class="btn-slot text x-huge" id="slot-elx-view-draft"></span>');
                var cfg = function (caption, icon, hint) {
                    return {cls: 'btn-toolbar x-huge icon-top', caption: caption, iconCls: 'toolbar__icon ' + icon, hint: hint, enableToggle: true, allowDepress: false, toggleGroup: 'elx-views',
                        lock: [Common.enumLock.lostConnect, Common.enumLock.disableOnStart], dataHint: '1', dataHintDirection: 'bottom', dataHintOffset: 'small'};
                };
                this.btnViewOutline = new Common.UI.Button(cfg(this.txtOutline, 'btn-big-contents', this.tipOutline));
                this.btnViewOutline.render($('#slot-elx-view-outline'));
                this.btnViewOutline.on('click', function () { me.setView('outline'); });
                this.btnViewDraft = new Common.UI.Button(cfg(this.txtDraft, 'btn-big-text', this.tipDraft));
                this.btnViewDraft.render($('#slot-elx-view-draft'));
                this.btnViewDraft.on('click', function () { me.setView('draft'); });
                v.lockedControls && v.lockedControls.push(this.btnViewOutline, this.btnViewDraft);
            }
            v.btnViewPrint.on('click', function () { me.setView('print'); });
            v.btnViewRead.on('click', function () { me.enterReadMode(); });
        },

        setView: function (view) {
            if (!this.api) return;
            if (this._state.read && view !== 'read') this.exitReadMode(view);
            else this.api.asc_SetWordView(view);
            Common.NotificationCenter.trigger('edit:complete');
        },

        command: function (cmd, value) {
            if (!this.api) return;
            this.api.asc_OutlineCommand(cmd, value);
            this.updateOutlineControls();
            Common.NotificationCenter.trigger('edit:complete');
        },

        onViewChanged: function (view) {
            var rt = this.getApplication().getController('RibbonTabs'), v = rt && rt.view, tb = this.toolbar, was = this._state.view;
            this._state.view = view;
            if (view !== 'read' && this._state.read) this.exitReadMode(view, true);
            if (view !== 'read') {
                v && v.btnViewPrint && v.btnViewPrint.toggle(view === 'print', true);
                this.btnViewOutline && this.btnViewOutline.toggle(view === 'outline', true);
                this.btnViewDraft && this.btnViewDraft.toggle(view === 'draft', true);
                v && v.btnViewRead && view !== 'print' && v.btnViewRead.toggle(false, true);
            }
            if (tb) {
                tb.setVisible('outlining', view === 'outline');
                if (view === 'outline' && was !== 'outline') {
                    tb.setTab('outlining');
                    this.updateOutlineControls();
                } else if (view !== 'outline' && was === 'outline' && tb.isTabActive('outlining')) {
                    tb.setTab('home');
                }
                tb.processPanelVisible && tb.processPanelVisible(null, true);
            }
        },

        updateOutlineControls: function () {
            var st = this.api && this.api.asc_GetOutlineState && this.api.asc_GetOutlineState();
            if (!st || !this.cmbLevel) return;
            this.cmbLevel.setValue(st.level);
            this.cmbShow.setValue(st.showLevel || 0);
            this.chkFormatting.setValue(!!st.showFormatting, true);
        },

        // ---- Read Mode ---------------------------------------------------------------------
        enterReadMode: function () {
            // RibbonTabs has just switched the document to viewing (doc:mode-changed came first): the reading
            // layout is applied now; leaving Read Mode goes back to the mode of before
            if (this._state.read) return;
            this._state.wantRead = true;
            this._state.modeBeforeRead = this._state.lastMode === 'view' ? this._state.prevMode : this._state.lastMode;
            if (this._state.lastMode === 'view' || !this.appOptions().isEdit) this.onDocModeChanged('view', true);
        },

        onDocModeChanged: function (mode, replay) {
            if (!replay) {
                this._state.prevMode = this._state.lastMode;
                this._state.lastMode = mode || 'edit';
            }
            if (mode === 'view' && this._state.wantRead && !this._state.read) {
                this._state.wantRead = false;
                this._state.read = true;
                var tc = this.getApplication().getController('Toolbar');
                this._state.folded = !!(this.toolbar && this.toolbar.isCompact && this.toolbar.isCompact());
                tc && tc.onChangeCompactView && !this._state.folded && tc.onChangeCompactView(null, true, true);
                this._state.zoom = (this.api.WordControl && this.api.WordControl.m_nZoomValue) || 100;
                // Word's Read Mode shows neither rulers nor spelling marks
                var logic = this.api.WordControl && this.api.WordControl.m_oLogicDocument;
                this._state.spell = !!(logic && logic.Spelling && logic.Spelling.Use);
                this._state.rulers = !!(this.api.WordControl && this.api.WordControl.m_bIsRuler);
                this._state.spell && this.api.asc_setSpellCheck(false);
                this._state.rulers && this.api.asc_SetViewRulers(false);
                var me = this;
                setTimeout(function () { me.applyReadLayout(); }, 50);
            } else if (mode !== 'view' && this._state.read) {
                this._state.modeBeforeRead = null;
                this.exitReadMode();
            }
            this._state.wantRead = false;
        },

        applyReadLayout: function () {
            var $area = $('#editor_sdk');
            if (!$area.length || !this.api) return;
            var mm = 25.4 / 96, w = $area.width() * mm, h = $area.height() * mm;
            var pageH = Math.max(80, h - 12), pageW = Math.max(80, Math.min(w - 20, pageH * 0.8, 200));
            this.api.asc_SetWordView('read', pageW, pageH);
            this.api.zoomFitToPage && this.api.zoomFitToPage();
            this.api.goToPage(0);
        },

        // view: the view to show; fromEngine: the engine already left the reading layout (Ctrl+Alt+P, O, N)
        exitReadMode: function (view, fromEngine) {
            if (!this._state.read) return;
            this._state.read = false;
            var tc = this.getApplication().getController('Toolbar');
            tc && tc.onChangeCompactView && !this._state.folded && tc.onChangeCompactView(null, false, true);
            !fromEngine && this.api && this.api.asc_SetWordView(view || 'print');
            this._state.spell && this.api.asc_setSpellCheck(true);
            this._state.rulers && this.api.asc_SetViewRulers(true);
            this.api && this.api.zoom && this.api.zoom(this._state.zoom || 100);
            var rt = this.getApplication().getController('RibbonTabs'), v = rt && rt.view;
            v && v.btnViewRead && v.btnViewRead.toggle(false, true);
            // back to the mode of before (Editing, or Reviewing...), not when the document was only viewed
            var back = this._state.modeBeforeRead;
            if (back && back !== 'view') Common.NotificationCenter.trigger('doc:mode-apply', back, true);
        },

        // keys of Read Mode: one screen (page) at a time, Home, End, a page number then Enter, Esc
        onReadKey: function (e) {
            if (!this._state.read || !this.api) return;
            if (Common.Utils.ModalWindow.isVisible() || $('.dropdown-menu:visible').length) return;
            var a = document.activeElement;
            if (a && /^(INPUT|TEXTAREA|SELECT)$/.test(a.tagName) && a.id !== 'area_id') return;
            var K = Common.UI.Keys, code = e.keyCode, mod = e.ctrlKey || e.altKey || e.metaKey;
            var count = this.api.getCountPages(), current = this.api.getCurrentPage(), target;
            if (code === K.ESC && !mod) {
                this.stop(e);
                this.exitReadMode();
                return;
            }
            if (mod) return;
            if (code >= 48 && code <= 57 && !e.shiftKey || code >= 96 && code <= 105) {
                this._state.pageInput += String((code >= 96 ? code - 96 : code - 48));
                this.stop(e);
                return;
            }
            if (code === K.RETURN && this._state.pageInput) {
                target = Math.max(0, Math.min(count - 1, parseInt(this._state.pageInput, 10) - 1));
                this._state.pageInput = '';
            } else if (code === K.HOME) target = 0;
            else if (code === K.END) target = count - 1;
            else if (code === K.PAGEDOWN || code === K.RIGHT || code === K.DOWN || code === K.SPACE && !e.shiftKey) target = Math.min(count - 1, current + 1);
            else if (code === K.PAGEUP || code === K.LEFT || code === K.UP || code === K.SPACE && e.shiftKey || code === K.BACKSPACE) target = Math.max(0, current - 1);
            if (target === undefined) return;
            this._state.pageInput = '';
            this.stop(e);
            this.api.goToPage(target);
        },

        stop: function (e) {
            e.preventDefault();
            e.stopPropagation();
        },

        // ---- field hints -------------------------------------------------------------------
        onFieldHint: function (kind, count) {
            var keys = this.getApplication().getController('WordEditKeys'),
                text = {locked: this.txtFieldLocked, unlocked: this.txtFieldUnlocked, nofield: this.txtNoField, lockedskip: this.txtLockedSkipped}[kind];
            if (keys && keys.showHint && text) keys.showHint(Common.Utils.String.format(text, count), 3000);
        },

        txtOutlining: 'Outlining',
        txtOutline: 'Outline',
        txtDraft: 'Draft',
        tipOutline: 'View the document as an outline where content shows as bulleted points',
        tipDraft: 'View the document as a draft for quick text editing (no headers, footers or page margins)',
        tipPromoteH1: 'Promote to Heading 1',
        tipPromote: 'Promote',
        tipDemote: 'Demote',
        tipDemoteBody: 'Demote to Body Text',
        tipMoveUp: 'Move Up',
        tipMoveDown: 'Move Down',
        tipExpand: 'Expand',
        tipCollapse: 'Collapse',
        tipOutlineLevel: 'Outline Level',
        txtLevel: 'Level {0}',
        txtBodyText: 'Body Text',
        txtAllLevels: 'All Levels',
        txtShowLevel: 'Show Level:',
        txtShowFormatting: 'Show Text Formatting',
        txtCloseOutline: 'Close Outline View',
        txtFieldLocked: 'Field locked: {0}',
        txtFieldUnlocked: 'Field unlocked: {0}',
        txtNoField: 'There is no field here',
        txtLockedSkipped: 'Locked fields are not updated: {0}'
    }, DE.Controllers.WordViews || {}));
});

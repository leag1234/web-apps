/*
 * (c) Copyright Ascensio System SIA 2010-2024
 *
 * This program is a free software product. You can redistribute it and/or
 * modify it under the terms of the GNU Affero General Public License (AGPL)
 * version 3 as published by the Free Software Foundation. In accordance with
 * Section 7(a) of the GNU AGPL its Section 15 shall be amended to the effect
 * that Ascensio System SIA expressly excludes the warranty of non-infringement
 * of any third-party rights.
 *
 * This program is distributed WITHOUT ANY WARRANTY; without even the implied
 * warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR  PURPOSE. For
 * details, see the GNU AGPL at: http://www.gnu.org/licenses/agpl-3.0.html
 *
 * The  interactive user interfaces in modified source and object code versions
 * of the Program must display Appropriate Legal Notices, as required under
 * Section 5 of the GNU AGPL version 3.
 *
 * All the Product's GUI elements, including illustrations and icon sets, as
 * well as technical writing content are licensed under the terms of the
 * Creative Commons Attribution-ShareAlike 4.0 International. See the License
 * terms at http://creativecommons.org/licenses/by-sa/4.0/legalcode
 *
 */

/**
 *  WordEditKeys.js
 *
 *  Interface of the Word editing shortcuts that the engine runs
 *  (sdkjs word/Editor/word-edit-keys.js): the status hint of F2 / Shift+F2
 *  (move or copy to a place) and of the Spike (Ctrl+F3), and the Paste
 *  Special dialog (Ctrl+Alt+V, Alt, H, V, S).
 */

define([
    'core',
    'documenteditor/main/app/view/PasteSpecialDialog'
], function () {
    'use strict';

    DE.Controllers.WordEditKeys = Backbone.Controller.extend(_.extend({
        initialize: function() {
            Common.NotificationCenter.on('app:ready', _.bind(this.onAppReady, this));
        },

        onAppReady: function(mode) {
            var me = this;
            this.mode = mode;
            this.api = this.getApplication().getController('Viewport').getApi();
            if (!this.api || !mode.isEdit) return;
            this.api.asc_registerCallback('asc_onMoveCopyMode', function(m) { me.onMoveCopyMode(m); });
            this.api.asc_registerCallback('asc_onSpikeChanged', function(count, added) { me.onSpikeChanged(count, added); });
            this.api.asc_registerCallback('asc_onPasteSpecialDialog', function() { me.openPasteSpecial(); });
            this.api.asc_registerCallback('asc_onDefineAutoText', function(text, data) { me.openDefineAutoText(text, data); });
            // Alt, H, V, then K, M, T or S: the letters of the Paste menu (the KeyTips of Word)
            document.addEventListener('keydown', _.bind(this.onPasteMenuKey, this), true);
            // keyboard in the print preview of the backstage (Print panel)
            document.addEventListener('keydown', _.bind(this.onPreviewKey, this), true);
            // Alt+F6, Home and End in the Comments pane
            document.addEventListener('keydown', _.bind(this.onWindowKey, this), true);
        },

        // ---------------------------------------------------------------- Alt+F6, comments list
        onWindowKey: function(e) {
            var K = Common.UI.Keys;
            if (e.keyCode === K.F6 && e.altKey && !e.ctrlKey && !e.shiftKey && !e.metaKey) {
                // back to the document from a pane or dialog that does not take the whole window
                // (a modal dialog keeps the focus until it is closed)
                if (Common.Utils.ModalWindow.isVisible()) return;
                e.preventDefault();
                e.stopPropagation();
                var a = document.activeElement;
                a && a !== document.body && a.blur && a.blur();
                this.api.asc_enableKeyEvents(true);
                Common.NotificationCenter.trigger('edit:complete');
                return;
            }
            // first and last comment of the Comments pane: Cmd+Home / Cmd+End (Ctrl elsewhere)
            if ((e.keyCode === K.HOME || e.keyCode === K.END) && (Common.Utils.isMac ? e.metaKey && !e.ctrlKey : e.ctrlKey && !e.metaKey) && !e.altKey && !e.shiftKey &&
                    $(document.activeElement).closest('#left-panel-comments').length && !/^(INPUT|TEXTAREA)$/.test(document.activeElement.tagName)) {
                var cc = this.getApplication().getController('Common.Controllers.Comments');
                if (!cc || !cc.collection) return;
                var list = _.sortBy(cc.collection.filter(function(m) { return !m.get('hide') && !m.get('removed') && m.get('position') >= 0; }), function(m) { return m.get('position'); });
                if (!list.length) return;
                e.preventDefault();
                e.stopPropagation();
                var id = (e.keyCode === K.HOME ? list[0] : list[list.length - 1]).get('uid');
                this.api.asc_selectComment(id);
                this.api.asc_showComment(id);
            }
        },

        // ---------------------------------------------------------------- print preview
        // The preview shows one page fitted in the window, so the arrow keys, Page Up and Page Down move
        // between pages; Ctrl+Home and Ctrl+End go to the first and the last page.
        onPreviewKey: function(e) {
            var print = this.getApplication().getController('Print');
            if (!print || !print._isPreviewVisible || !print._navigationPreview || !$('#print-preview-box:visible').length) return;
            var K = Common.UI.Keys, nav = print._navigationPreview, active = document.activeElement,
                current = nav.currentPreviewPage, last = Math.max(0, nav.pageCount - 1), target;
            if (e.altKey || e.metaKey || e.shiftKey) return;
            // text boxes, combo lists and menus keep their keys
            if (active && ((/^(INPUT|TEXTAREA|SELECT)$/.test(active.tagName) && active.id !== 'area_id') || $(active).closest('.dropdown-menu, .combo-dataview, .menu-picker-container').length) || $('.dropdown-menu:visible').length) return;
            if (e.ctrlKey) {
                if (e.keyCode === K.HOME) target = 0;
                else if (e.keyCode === K.END) target = last;
            } else if (e.keyCode === K.PAGEDOWN || e.keyCode === K.RIGHT || e.keyCode === K.DOWN) target = Math.min(current + 1, last);
            else if (e.keyCode === K.PAGEUP || e.keyCode === K.LEFT || e.keyCode === K.UP) target = Math.max(current - 1, 0);
            if (target === undefined) return;
            e.preventDefault();
            e.stopPropagation();
            target !== current && this.api.goToPage(target);
        },

        onPasteMenuKey: function(e) {
            if (e.ctrlKey || e.altKey || e.metaKey || e.shiftKey || !$('.word-paste-menu:visible').length) return;
            var P = Asc.c_oSpecialPasteProps, fmt = this.getApplication().getController('WordFormatting'),
                keys = {K: P.sourceformatting, M: P.mergeFormatting, T: P.keepTextOnly, S: 'special'},
                value = keys[String.fromCharCode(e.keyCode)];
            if (value === undefined) return;
            e.preventDefault();
            e.stopPropagation();
            Common.UI.Menu.Manager.hideAll();
            if (value === 'special') this.openPasteSpecial();
            else fmt && fmt.pasteWithOption(value);
        },

        // ---------------------------------------------------------------- status hint
        // a small message above the status bar (Word writes these in its status bar)
        showHint: function(text, timeout) {
            var $hint = $('#eo-status-hint');
            if (!$hint.length) {
                $('<style id="eo-status-hint-style"></style>').text(
                    '#eo-status-hint{position:fixed;left:12px;bottom:34px;z-index:20000;max-width:60vw;padding:6px 12px;border-radius:4px;' +
                    'background:var(--background-toolbar,#fff);color:var(--text-normal,#333);border:1px solid var(--border-regular-control,#ccc);' +
                    'box-shadow:0 2px 8px rgba(0,0,0,.2);font-size:12px;pointer-events:none;}'
                ).appendTo('head');
                $hint = $('<div id="eo-status-hint" role="status" aria-live="polite"></div>').appendTo(document.body);
            }
            clearTimeout(this._hintTimer);
            $hint.text(text).show();
            if (timeout) this._hintTimer = setTimeout(_.bind(this.hideHint, this), timeout);
        },

        hideHint: function() {
            clearTimeout(this._hintTimer);
            $('#eo-status-hint').hide().text('');
        },

        // ---------------------------------------------------------------- F2, Shift+F2
        onMoveCopyMode: function(m) {
            if (m) this.showHint(m === 2 ? this.textCopyTo : this.textMoveTo);
            else this.hideHint();
        },

        // ---------------------------------------------------------------- Spike
        onSpikeChanged: function(count, added) {
            this.showHint(count ? Common.Utils.String.format(this.textSpike, count) : this.textSpikeEmpty, 2500);
        },

        // ---------------------------------------------------------------- Paste Special
        openPasteSpecial: function() {
            var me = this;
            if (Common.Utils.ModalWindow.isVisible() || !this.mode || !this.mode.isEdit) return;
            Common.UI.Menu.Manager.hideAll();
            (new DE.Views.PasteSpecialDialog({
                handler: function(result, value) {
                    if (result === 'ok') {
                        var fmt = me.getApplication().getController('WordFormatting');
                        fmt && fmt.pasteWithOption && fmt.pasteWithOption(value);
                    }
                    me.api.asc_enableKeyEvents(true);
                    Common.NotificationCenter.trigger('edit:complete');
                }
            })).show();
        },

        // ---------------------------------------------------------------- AutoText (Alt+F3, name + Enter)
        // The blocks are kept per user in this browser (localStorage); the engine inserts a block when Enter is
        // pressed right after its name (sdkjs word/Editor/word-edit-keys.js).
        openDefineAutoText: function(text, data) {
            var me = this;
            if (Common.Utils.ModalWindow.isVisible() || !this.mode || !this.mode.isEdit) return;
            (new Common.Views.TextInputDialog({
                width: 400,
                title: this.textAutoTextTitle,
                label: this.textAutoTextName,
                description: this.textAutoTextHelp,
                value: (text || '').trim().split(/\s+/)[0].substr(0, 24),
                inputConfig: {
                    allowBlank: false,
                    validation: function(value) { return /\s/.test(value) ? me.textAutoTextNoSpace : true; }
                },
                handler: function(result, value) {
                    result === 'ok' && me.saveAutoText(value, text, data);
                    me.api.asc_enableKeyEvents(true);
                    Common.NotificationCenter.trigger('edit:complete');
                }
            })).show();
        },

        saveAutoText: function(name, text, data) {
            try {
                var names = JSON.parse(window.localStorage.getItem('eo-autotext-index') || '[]');
                window.localStorage.setItem('eo-autotext:' + name, JSON.stringify({data: data, text: text}));
                names.indexOf(name) < 0 && names.push(name);
                window.localStorage.setItem('eo-autotext-index', JSON.stringify(names));
                this.showHint(Common.Utils.String.format(this.textAutoTextSaved, name), 5000);
            } catch (e) {
                this.showHint(this.textAutoTextFailed, 5000);
            }
        },

        textMoveTo: 'Move to where? Press Enter',
        textCopyTo: 'Copy to where? Press Enter',
        textSpike: 'Spike: {0}',
        textSpikeEmpty: 'The Spike is empty',
        textAutoTextTitle: 'Create AutoText',
        textAutoTextName: 'Name:',
        textAutoTextHelp: 'To insert it, type the name and press Enter.',
        textAutoTextNoSpace: 'The name cannot contain spaces',
        textAutoTextSaved: 'AutoText "{0}" saved. Type its name and press Enter to insert it',
        textAutoTextFailed: 'The AutoText could not be saved in this browser'
    }, DE.Controllers.WordEditKeys || {}));
});

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

    DE.Controllers.WordEditKeys = Backbone.Controller.extend({
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
            // Alt, H, V, then K, M, T or S: the letters of the Paste menu (the KeyTips of Word)
            document.addEventListener('keydown', _.bind(this.onPasteMenuKey, this), true);
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

        textMoveTo: 'Move to where? Press Enter',
        textCopyTo: 'Copy to where? Press Enter',
        textSpike: 'Spike: {0}',
        textSpikeEmpty: 'The Spike is empty'
    });
});

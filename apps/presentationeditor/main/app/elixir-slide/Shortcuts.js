/*
 * SPDX-FileCopyrightText: 2026 ElixirOffice contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 *
 * PE.Controllers.SlideShortcuts: the PowerPoint keyboard shortcuts that the
 * presentation editor did not have (matrix: docs/shortcuts-powerpoint.csv in
 * the ElixirOffice repository, from Microsoft's public "Keyboard shortcuts in
 * PowerPoint" page).
 *
 * A capture-phase keydown listener runs before the editor's own handlers. A
 * shortcut is either run directly (an existing controller or API call) or
 * remapped to the stock shortcut of the same command (synthetic keydown with
 * the stock keys), so the engine's own behaviour is reused. Ctrl means Cmd on
 * macOS, like the rest of the editor.
 */
define([
    'core'
], function () {
    'use strict';

    var K = { F3: 114, F5: 116, F9: 120, EQUAL: 187, MINUS: 189, PERIOD: 190, COMMA: 188, BR_L: 219, BR_R: 221,
              LEFT: 37, RIGHT: 39, UP: 38, DOWN: 40, C: 67, D: 68, F: 70, M: 77, O: 79, V: 86 };
    var isMac = Common.Utils.isMac;

    PE.Controllers.SlideShortcuts = Backbone.Controller.extend({

        initialize: function () {
            Common.NotificationCenter.on('app:ready', this.onAppReady.bind(this));
        },

        app: function (n) { return this.getApplication().getController(n); },

        onAppReady: function (mode) {
            if (this.ready || !mode || !mode.isEdit) return;
            this.ready = true;
            this.api = this.app('Viewport').getApi();
            this.tb = this.app('Toolbar').toolbar;
            var me = this;
            window.addEventListener('keydown', function (e) { me.onKeyDown(e); }, true);
        },

        // where the key goes: no shortcut while a dialog, a menu or a form field has the focus
        inEditor: function (e) {
            var t = e.target;
            if ($('.asc-window.modal:visible').length) return false;
            if (t && /^(INPUT|SELECT)$/.test(t.tagName) && !$(t).closest('#editor_sdk').length) return false;
            if (t && t.tagName === 'TEXTAREA' && !$(t).closest('#editor_sdk').length) return false;
            return true;
        },

        previewOpen: function () { return $('#pe-preview').is(':visible'); },

        onKeyDown: function (e) {
            var ctrl = isMac ? e.metaKey : e.ctrlKey, k = e.keyCode, sh = e.shiftKey, alt = e.altKey,
                otherCtrl = isMac ? e.ctrlKey : e.metaKey;
            if (otherCtrl) return;
            if (this.previewOpen()) { this.onShowKey(e, ctrl, sh, alt); return; }
            if (!this.inEditor(e)) return;
            var act = this.match(k, ctrl, sh, alt, e);
            if (!act) return;
            e.preventDefault();
            e.stopImmediatePropagation();
            act.call(this, e);
        },

        match: function (k, ctrl, sh, alt, e) {
            var me = this;
            // ---- slide show (F5, Shift+F5, Alt+F5)
            if (k === K.F5 && !ctrl && !alt) return function () { me.show(sh); };
            if (k === K.F5 && alt && !ctrl && !sh) return function () { me.presenter(); };
            // ---- format text
            // Ctrl+= subscript, Ctrl+Shift++ superscript while editing text; zoom in otherwise, as in PowerPoint
            var inText = this.tb.btnSubscript && !this.tb.btnSubscript.isDisabled();
            if (ctrl && !alt && k === K.EQUAL && inText) return function (ev) { me.remap(ev, sh ? K.COMMA : K.PERIOD, true, false, false); };
            if (ctrl && !alt && !sh && (k === 107 || k === 109)) return function () { k === 107 ? me.api.zoomIn() : me.api.zoomOut(); };
            if (ctrl && sh && !alt && k === K.PERIOD) return function () { me.click(me.tb.btnIncFontSize); };
            if (ctrl && sh && !alt && k === K.COMMA) return function () { me.click(me.tb.btnDecFontSize); };
            // Ctrl+Shift+C / Ctrl+Shift+V: copy / paste formatting (the engine's Ctrl+Alt+C / Ctrl+Alt+V commands)
            if (ctrl && sh && !alt && (k === K.C || k === K.V)) return function () { me.formatPainter(k === K.C); };
            if (ctrl && sh && !alt && k === K.F) return function () { var dh = me.app('DocumentHolder'); dh && dh.onParagraphAdvanced && dh.onParagraphAdvanced(); };
            if (sh && !ctrl && !alt && k === K.F3) return function () { me.cycleCase(); };
            if (alt && sh && !ctrl && k === K.LEFT) return function () { me.click(me.tb.btnDecLeftOffset); };   // promote
            if (alt && sh && !ctrl && k === K.RIGHT) return function () { me.click(me.tb.btnIncLeftOffset); };  // demote
            // ---- find (Ctrl+F: the Find pane with the selected text, like PowerPoint), hyperlink
            if (ctrl && !sh && !alt && k === K.F) return function () {
                var c = me.app('ElixirChrome'), sel = me.api.asc_GetSelectedText ? $.trim(me.api.asc_GetSelectedText() || '') : '';
                c && c.findInDocument(sel);
            };
            if (ctrl && !sh && !alt && k === 75 && this.tb.btnInsertHyperlink && !this.tb.btnInsertHyperlink.isDisabled())
                return function () { me.click(me.tb.btnInsertHyperlink); };
            // ---- slides and objects
            if (ctrl && sh && !alt && k === K.D) return function () { me.tb.fireEvent('duplicate:slide'); };
            if (ctrl && alt && !sh && k === K.M) return function () { var dh = me.app('DocumentHolder'); dh && dh.addComment(); };
            if (ctrl && sh && !alt && k === K.BR_R) return function () { me.api.shapes_bringForward(); me.done(); };
            if (ctrl && sh && !alt && k === K.BR_L) return function () { me.api.shapes_bringBackward(); me.done(); };
            if (alt && !ctrl && !sh && (k === K.LEFT || k === K.RIGHT) && this.hasObject()) return function () { me.rotate(k === K.RIGHT ? 15 : -15); };
            if (alt && !ctrl && !sh && k === K.EQUAL) return function () { me.click(me.tb.btnInsertEquation); };
            // ---- views
            if (ctrl && alt && !sh && k === K.O) return function () { me.api.zoomFitToPage(); };
            if (ctrl && !alt && !sh && k === 112) return function () {   // Ctrl+F1: collapse / pin the ribbon
                var tc = me.app('Toolbar');
                tc && tc.onChangeCompactView && tc.onChangeCompactView(null, !me.tb.isCompact());
            };
            if (alt && !ctrl && !sh && k === K.F9) return function () { me.api.asc_setShowGuides(!me.api.asc_getShowGuides()); me.done(); };
            if (sh && !ctrl && !alt && k === K.F9) return function () { me.api.asc_setShowGridlines(!me.api.asc_getShowGridlines()); me.done(); };
            return null;
        },

        // ---- delivering a presentation (PowerPoint keys the slide show did not have) ------------
        onShowKey: function (e, ctrl, sh, alt) {
            var k = e.keyCode, api = this.api, me = this, handled = true;
            if (ctrl || alt) return;
            var blank = $('#elx-show-blank');
            if (blank.length && k !== 66 && k !== 87 && k !== 190 && k !== 188) { blank.remove(); }
            if (k >= 48 && k <= 57 || k >= 96 && k <= 105) {               // slide number, then Enter
                this._goto = (this._goto || '') + String((k >= 96 ? k - 96 : k - 48));
            } else if (k === 13 && this._goto) {
                api.DemonstrationGoToSlide(Math.max(0, parseInt(this._goto, 10) - 1));
                this._goto = '';
            } else if (k === 78 && !sh) { api.DemonstrationNextSlide(); }        // N
            else if ((k === 80 && !sh) || k === 8) { api.DemonstrationPrevSlide(); } // P, Backspace
            else if (k === 66 || k === 190) { this.blank('#000'); }             // B, period: black screen
            else if (k === 87 || k === 188) { this.blank('#fff'); }             // W, comma: white screen
            else handled = false;
            if (!(k >= 48 && k <= 57 || k >= 96 && k <= 105) && k !== 13) this._goto = '';
            if (handled) { e.preventDefault(); e.stopImmediatePropagation(); }
            void me;
        },

        blank: function (color) {
            var $b = $('#elx-show-blank');
            if ($b.length && $b.data('color') === color) { $b.remove(); return; }
            $b.remove();
            $('<div id="elx-show-blank"></div>').data('color', color)
                .css({ position: 'absolute', left: 0, top: 0, right: 0, bottom: 0, background: color, 'z-index': 1000, cursor: 'none' })
                .on('click', function () { $(this).remove(); })
                .appendTo('#pe-preview');
        },

        done: function () { Common.NotificationCenter.trigger('edit:complete', this.tb); },

        click: function (btn) {
            if (!btn || btn.isDisabled()) return;
            var $c = btn.cmpEl || btn.$el;
            var $b = $c && ($c.is('button') ? $c : $c.find('button').first());
            $b && $b[0] && $b[0].click();
        },

        // replay the stock shortcut of the same command on the element that got the key
        remap: function (e, keyCode, ctrl, shift, alt) {
            var ev = new KeyboardEvent('keydown', { bubbles: true, cancelable: true, ctrlKey: !isMac && ctrl, metaKey: isMac && ctrl,
                shiftKey: shift, altKey: alt });
            Object.defineProperty(ev, 'keyCode', { get: function () { return keyCode; } });
            Object.defineProperty(ev, 'which', { get: function () { return keyCode; } });
            this._remapping = true;
            (e.target || document).dispatchEvent(ev);
            this._remapping = false;
        },

        formatPainter: function (copy) {
            var doc = this.api.WordControl && this.api.WordControl.m_oLogicDocument;
            if (!doc) return;
            if (copy) doc.Document_Format_Copy();
            else if (doc.CanEdit()) doc.Document_Format_Paste();
        },

        hasObject: function () {
            var T = Asc.c_oAscTypeSelectElement, has = false;
            _.each(this.api.getSelectedElements() || [], function (o) {
                var t = o.get_ObjectType();
                if (t === T.Shape || t === T.Image || t === T.Chart) has = true;
            });
            return has;
        },

        rotate: function (deg) {
            var p = new Asc.asc_CShapeProperty();
            p.asc_putRotAdd(deg * Math.PI / 180);
            this.api.ShapeApply(p);
            this.done();
        },

        // Shift+F3: lowercase, UPPERCASE, Capitalize Each Word, like PowerPoint
        cycleCase: function () {
            var T = Asc.c_oAscChangeTextCaseType, seq = [T.LowerCase, T.UpperCase, T.CapitalizeWords];
            this._case = ((this._case === undefined ? -1 : this._case) + 1) % seq.length;
            this.api.asc_ChangeTextCase(seq[this._case]);
        },

        show: function (current) {
            var r = this.app('SlideRibbon');
            r && r.startShow ? r.startShow(!!current) : Common.NotificationCenter.trigger('preview:start', current ? this.api.getCurrentPage() : 0);
        },

        presenter: function () {
            var tbc = this.app('Toolbar');
            tbc && tbc.onPreview(0, true, false);
        }
    });
});

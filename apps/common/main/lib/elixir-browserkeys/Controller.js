/*
 * SPDX-FileCopyrightText: 2026 ElixirOffice contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 *
 * Common.Controllers.ElixirBrowserKeys: Office shortcuts that the browser keeps for itself (docs/shortcuts.md, "Browser-reserved keys").
 *  - Office-compatible alternatives that do reach the page (Ctrl+Alt+Shift+W/T/N/Tab). The alternative plays the reserved
 *    key in the editor: it is turned back into the original key event, so the editor needs no second implementation.
 *    Spreadsheet sheet navigation has its own keys in ExcelShortcuts (Ctrl+Alt+Page Down / Page Up).
 *  - a one-time hint after a combination that the page cannot see but can detect afterwards (Ctrl held, nothing typed,
 *    the window lost the focus within 2.5 s);
 *  - a "Keyboard shortcuts in the browser" dialog: reserved keys, what to press instead, Install app, Full screen.
 *  - the top-level page (web-apps/apps/api/documents/api.js) owns Keyboard Lock, the install prompt and full screen; this
 *    controller talks to it with window.postMessage ({elixirBrowserKeys: ...} and {elixirBrowserKeysStatus: ...}).
 */

define(['core'], function () {
    'use strict';

    var BK = window.Common.ElixirBrowserKeys = window.Common.ElixirBrowserKeys || {};
    var MAC = /Mac|iPhone|iPad/.test(navigator.platform || '');
    var KEYCODE = {KeyW: 87, KeyN: 78, KeyT: 84, KeyF: 70, Tab: 9};
    var HINT_KEY = 'elixir-browserkeys-hint-seen';
    BK.isLocal = function () {
        try { if (window.parent.eoDesktop && typeof window.parent.eoDesktop.onAction === 'function') return true; } catch (e) { /* cross-origin hosted parent */ }
        var app = window.DE || window.SSE || window.PE || window.PDFE;
        var main = app && app.getController('Main');
        return !!(main && (main.appOptions && main.appOptions.isElixirLocal || main.api && main.api.eoLocalMode));
    };


    function editor() { return window.DE ? 'word' : window.SSE ? 'excel' : window.PE ? 'powerpoint' : ''; }
    function esc(s) { return String(s).replace(/[&<>"]/g, function (c) { return {'&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;'}[c]; }); }
    function store(k, v) { try { if (v === undefined) return window.localStorage.getItem(k); window.localStorage.setItem(k, v); } catch (e) { return BK['_mem' + k]; } }

    // Alternatives that reach the page. code: physical key; to: the key event played in the editor (cmd = Ctrl, Cmd on macOS)
    // Pressed as Ctrl+Alt+Shift+<code> (macOS: Cmd+Option+Shift+<code>). Chosen free in the three shortcut matrices.
    var ALTS = {
        word: [
            {code: 'KeyW', to: {code: 'KeyW', shift: true}, key: 'W'},         // Ctrl+Shift+W underline words only
            {code: 'KeyT', to: {code: 'KeyT'}, key: 'T'},                        // Ctrl+T hanging indent
            {code: 'KeyN', to: {code: 'KeyN', shift: true}, key: 'N'},         // Ctrl+Shift+N Normal style
            {code: 'Tab', to: {code: 'Tab'}, key: 'Tab', noMac: true}          // Ctrl+Tab tab character in a table cell
        ],
        excel: [
            {code: 'KeyT', to: {code: 'KeyT'}, key: 'T'},                        // Ctrl+T create table
            {code: 'KeyN', action: 'new', key: 'N'}                             // Ctrl+N new workbook
        ],
        powerpoint: [
            {code: 'KeyT', to: {code: 'KeyF', shift: true}, key: 'T'},         // Ctrl+T Font dialog
            {code: 'KeyN', action: 'new', key: 'N'},                            // Ctrl+N new presentation
            {code: 'Tab', to: {code: 'Tab'}, key: 'Tab', noMac: true}          // Ctrl+Tab tab in a table cell
        ]
    };

    // rows of the dialog: [what the browser does with the key, [Windows/Linux key, macOS key], [use instead (Windows), (macOS)], editors]
    var ROWS = [
        ['Ctrl+PageDown, Ctrl+PageUp', 'Next, previous sheet', 'Ctrl+Alt+PageDown, Ctrl+Alt+PageUp', 'Option+Right, Option+Left', ['excel']],
        ['Ctrl+Shift+PageDown, Ctrl+Shift+PageUp', 'Select next, previous sheet', 'Ctrl+Alt+Shift+PageDown, Ctrl+Alt+Shift+PageUp', 'Ctrl+Option+Shift+PageDown, Ctrl+Option+Shift+PageUp', ['excel']],
        ['Ctrl+PageDown, Ctrl+PageUp', 'Next, previous page', 'Ctrl+G (Go To page)', 'Cmd+PageDown works', ['word']],
        ['Ctrl+N', 'New workbook', 'Ctrl+Alt+Shift+N, File > New', 'Cmd+Option+Shift+N, File > New', ['excel']],
        ['Ctrl+N', 'New presentation', 'Ctrl+Alt+Shift+N, File > New', 'Cmd+Option+Shift+N, File > New', ['powerpoint']],
        ['Ctrl+N', 'New document', 'File > New (Alt, F, N)', 'File > New', ['word']],
        ['Ctrl+W', 'Close', 'File > Close (unsaved changes are asked first)', 'File > Close', ['word', 'excel', 'powerpoint']],
        ['Ctrl+T', 'Create table', 'Ctrl+L, Ctrl+Alt+Shift+T', 'Cmd+L, Cmd+Option+Shift+T', ['excel']],
        ['Ctrl+T', 'Hanging indent', 'Ctrl+Alt+Shift+T', 'Cmd+Option+Shift+T', ['word']],
        ['Ctrl+T', 'Font dialog', 'Ctrl+Shift+F, Ctrl+Alt+Shift+T', 'Cmd+Shift+F, Cmd+Option+Shift+T', ['powerpoint']],
        ['Ctrl+Shift+W', 'Underline words only', 'Ctrl+Alt+Shift+W', 'Cmd+Option+Shift+W', ['word']],
        ['Ctrl+Shift+N', 'Normal style', 'Ctrl+Alt+Shift+N', 'Cmd+Option+Shift+N', ['word']],
        ['Ctrl+Tab', 'Tab character in a table cell', 'Ctrl+Alt+Shift+Tab', 'Option+Tab', ['word']],
        ['Ctrl+Tab', 'Tab in a table cell', 'Ctrl+Alt+Shift+Tab', 'Option+Tab', ['powerpoint']]
    ];

    var CSS = [
        '.elx-bk-bar{position:fixed;z-index:1001;left:50%;bottom:44px;transform:translateX(-50%);box-sizing:border-box;width:680px;max-width:calc(100vw - 32px);',
        'display:flex;align-items:center;gap:12px;padding:10px 12px 10px 16px;background:var(--background-normal,#fff);color:var(--text-normal,#222);',
        'border:1px solid var(--border-regular-control,#ccc);border-radius:6px;box-shadow:0 4px 18px rgba(0,0,0,.22);font:12px/1.4 var(--font-family-base-custom,Arial,sans-serif);}',
        '.elx-bk-msg{flex:1 1 auto;min-width:0;}',
        '.elx-bk-btn{font:inherit;padding:5px 12px;border-radius:4px;cursor:pointer;flex:0 0 auto;border:1px solid var(--background-accent-button,#446995);background:var(--background-accent-button,#446995);color:var(--text-inverse,#fff);}',
        '.elx-bk-x{font:inherit;font-size:16px;line-height:1;padding:2px 6px;border:0;background:none;cursor:pointer;color:inherit;flex:0 0 auto;}',
        '.elx-bk-tbl{border-collapse:collapse;width:100%;font-size:12px;}.elx-bk-tbl th,.elx-bk-tbl td{text-align:left;padding:4px 8px;border-bottom:1px solid var(--border-regular-control,#ddd);vertical-align:top;}',
        '.elx-bk-note{margin:10px 0 0;line-height:1.5;}.elx-bk-act{display:flex;gap:8px;margin-top:12px;}.elx-bk-act button{font:inherit;padding:5px 12px;cursor:pointer;}'
    ].join('');

    Common.Controllers.ElixirBrowserKeys = Backbone.Controller.extend(_.extend({
        status: {},

        initialize: function () {
            var me = this;
            BK.controller = this;
            BK.alts = ALTS;
            me.ed = editor();
            // before any other key listener of the editor (this one is registered when the controller is created)
            window.addEventListener('keydown', me.onKeyDown.bind(me), true);
            window.addEventListener('message', me.onMessage.bind(me), false);
            me.trackHint();
            Common.NotificationCenter.on('app:ready', me.onAppReady.bind(me));
            Common.NotificationCenter.on('elixir:browserkeys:help', function () { me.showDialog(); });
            me.askStatus();
        },

        onLaunch: function () {},

        onAppReady: function () {
            if (BK.isLocal() || this.styled) return;
            this.styled = true;
            var st = document.createElement('style');
            st.textContent = CSS;
            document.head.appendChild(st);
        },

        // ---- alternatives ---------------------------------------------------------------------
        onKeyDown: function (e) {
            if (e.isTrusted === false) return;
            var mod = MAC ? (e.metaKey && !e.ctrlKey) : (e.ctrlKey && !e.metaKey);
            if (!mod) return;
            // Ctrl+N reaches the page only in an installed app or in full screen: File > New in the spreadsheet and presentation editors
            if (e.code === 'KeyN' && !e.shiftKey && !e.altKey && (this.ed === 'excel' || this.ed === 'powerpoint') && !Common.Utils.ModalWindow.isVisible()) {
                e.preventDefault(); e.stopImmediatePropagation();
                return this.openNew();
            }
            if (!e.shiftKey || !e.altKey) return;
            var alt = ALTS[this.ed] && ALTS[this.ed].filter(function (a) { return a.code === e.code && !(MAC && a.noMac); })[0];
            if (!alt) return;
            e.preventDefault();
            e.stopImmediatePropagation();
            if (alt.action === 'new') return this.openNew();
            this.play(alt.to);
        },

        // plays the original key (Ctrl or Cmd, optional Shift) in the focused element, like the keyboard would
        play: function (to) {
            var target = document.activeElement || document.body;
            var init = {key: to.code === 'Tab' ? 'Tab' : to.code.slice(3).toLowerCase(), code: to.code, ctrlKey: !MAC, metaKey: MAC, shiftKey: !!to.shift, altKey: false,
                bubbles: true, cancelable: true, composed: true, keyCode: KEYCODE[to.code], which: KEYCODE[to.code]};
            ['keydown', 'keyup'].forEach(function (type) {
                var ev = new KeyboardEvent(type, init);
                try { Object.defineProperty(ev, 'keyCode', {get: function () { return KEYCODE[to.code]; }}); Object.defineProperty(ev, 'which', {get: function () { return KEYCODE[to.code]; }}); } catch (x) { /* read-only: the init values are used */ }
                target.dispatchEvent(ev);
            });
        },

        openNew: function () {
            try {
                var lm = this.getApplication().getController('LeftMenu');
                Common.UI.Menu.Manager.hideAll();
                lm.leftMenu.showMenu('file');
                lm.leftMenu.menuFile.selectMenu('new');
            } catch (e) { console.warn('ElixirBrowserKeys: File > New', e); }
        },

        // ---- the top-level page (Keyboard Lock, install, full screen) -----------------------------
        askStatus: function () { if (!BK.isLocal()) this.toTop({ask: 'status'}); },

        toTop: function (msg) { try { window.parent !== window && window.parent.postMessage({elixirBrowserKeys: msg}, '*'); } catch (e) { /* no parent */ } },

        onMessage: function (e) {
            if (e.source !== window.parent || !e.data || !e.data.elixirBrowserKeysStatus) return;
            this.status = e.data.elixirBrowserKeysStatus;
            BK.status = this.status;
            this.dlgSync && this.dlgSync();
        },

        // ---- hint -----------------------------------------------------------------------------
        // The keys the browser keeps never reach the page, but their effect is visible: Ctrl was pressed alone and the window
        // lost the focus right away (tab switch, new window). Shown once, when the user comes back.
        trackHint: function () {
            if (BK.isLocal()) return;
            var me = this, ctrl = null, pending = false;
            var isMod = function (k) { return k === 'Control' || k === 'Meta' || k === 'Shift' || k === 'Alt'; };
            window.addEventListener('keydown', function (e) {
                if (e.isTrusted === false) return;
                if (e.key === 'Control' || e.key === 'Meta') { if (!e.repeat) ctrl = {t: Date.now(), other: false}; }
                else if (ctrl && !isMod(e.key)) ctrl.other = true;
            }, true);
            window.addEventListener('keyup', function (e) { if (e.key === 'Control' || e.key === 'Meta') ctrl = null; }, true);
            var leave = function () {
                if (ctrl && !ctrl.other && Date.now() - ctrl.t < 2500 && !me.status.standalone && !me.status.fullscreen) pending = true;
                ctrl = null;
            };
            var back = function () {
                ctrl = null;
                if (pending) { pending = false; setTimeout(function () { me.showHint(); }, 500); }
            };
            window.addEventListener('blur', leave);
            document.addEventListener('visibilitychange', function () { document.hidden ? leave() : back(); });
            window.addEventListener('focus', back);
            BK.simulateLeave = function () { ctrl = {t: Date.now(), other: false}; leave(); back(); };   // test hook
        },

        showHint: function (force) {
            var me = this;
            if (BK.isLocal()) return;
            if (!force && (store(HINT_KEY) || document.getElementById('elx-bk-hint'))) return;
            store(HINT_KEY, '1');
            BK._mem = 1;
            var bar = document.createElement('div');
            bar.id = 'elx-bk-hint';
            bar.className = 'elx-bk-bar';
            bar.setAttribute('role', 'status');
            bar.innerHTML = '<span class="elx-bk-msg">' + esc(this.txtHint) + '</span><button class="elx-bk-btn" type="button">' + esc(this.txtSee) +
                '</button><button class="elx-bk-x" type="button" aria-label="' + esc(this.txtDismiss) + '" title="' + esc(this.txtDismiss) + '">×</button>';
            document.body.appendChild(bar);
            var close = function () { bar.parentNode && bar.parentNode.removeChild(bar); };
            bar.querySelector('.elx-bk-x').addEventListener('click', close);
            bar.querySelector('.elx-bk-btn').addEventListener('click', function () { close(); me.showDialog(); });
            setTimeout(close, 20000);
        },

        // ---- dialog ---------------------------------------------------------------------------
        showDialog: function () {
            var me = this;
            if (BK.isLocal()) return;
            if (this.dlg && this.dlg.isVisible && this.dlg.isVisible()) return;
            var rows = ROWS.filter(function (r) { return r[4].indexOf(me.ed) >= 0; }).map(function (r) {
                var mac = r[0].replace(/Ctrl/g, 'Cmd');
                return '<tr><td>' + esc(MAC ? mac : r[0]) + '</td><td>' + esc(r[1]) + '</td><td>' + esc(MAC ? r[3] : r[2]) + '</td></tr>';
            }).join('');
            this.askStatus();
            var win = this.dlg = new Common.UI.Window({
                title: this.txtTitle, width: 640, cls: 'modal-dlg elx-bk-dlg',
                tpl: '<div class="box" style="padding: 8px 4px; max-height: 70vh; overflow: auto;">' +
                    '<p class="elx-bk-note" style="margin-top:0">' + esc(this.txtIntro) + '</p>' +
                    '<table class="elx-bk-tbl"><thead><tr><th>' + esc(this.txtKey) + '</th><th>' + esc(this.txtAction) + '</th><th>' + esc(this.txtInstead) + '</th></tr></thead><tbody>' + rows + '</tbody></table>' +
                    '<p class="elx-bk-note">' + esc(this.txtWays) + '</p>' +
                    '<div class="elx-bk-act"><button type="button" class="btn normal elx-bk-install" style="display:none">' + esc(this.txtInstall) + '</button>' +
                    '<button type="button" class="btn normal elx-bk-full" style="display:none">' + esc(this.txtFull) + '</button></div>' +
                    '<p class="elx-bk-note elx-bk-state"></p>' +
                    '</div><div class="footer center"><button class="btn normal dlg-btn primary" result="ok">' + esc(this.txtClose) + '</button></div>'
            });
            win.on('render:after', function () {
                var $w = win.$window, $i = $w.find('.elx-bk-install'), $f = $w.find('.elx-bk-full'), $s = $w.find('.elx-bk-state');
                me.dlgSync = function () {
                    var st = me.status || {};
                    $i.toggle(!!st.canInstall && !st.standalone);
                    $f.toggle(!!st.canFullscreen && !st.fullscreen);
                    $s.text(st.standalone ? me.txtStandalone : st.fullscreen ? me.txtFullscreen : '');
                };
                me.dlgSync();
                $i.on('click', function () { me.toTop({do: 'install'}); });
                $f.on('click', function () { me.toTop({do: 'fullscreen'}); win.close(); });
                $w.find('.dlg-btn').on('click', function () { win.close(); });
                win.on('close', function () { me.dlgSync = null; });
            });
            win.show();
        },

        txtTitle: 'Keyboard shortcuts in the browser',
        txtIntro: 'Browsers keep some shortcuts for themselves (tabs and windows), so a web page never receives them. Use the alternatives below.',
        txtKey: 'Browser key', txtAction: 'In Office', txtInstead: 'Use instead',
        txtWays: 'For the Office keys themselves: install ElixirOffice as an app (its window has no tabs, so Ctrl+Tab, Ctrl+PageDown and Ctrl+T reach the editor), or use full screen (Chromium browsers then hand most of these keys to the page; hold Esc to leave).',
        txtInstall: 'Install ElixirOffice as an app', txtFull: 'Full screen',
        txtStandalone: 'You are using the installed app.', txtFullscreen: 'Full screen: the keys above are sent to the editor.',
        txtClose: 'Close',
        txtHint: 'The browser kept that shortcut (for example Ctrl+PageDown, Ctrl+N or Ctrl+W). ElixirOffice has alternatives, and the installed app or full screen receives most of them.',
        txtSee: 'Show shortcuts', txtDismiss: 'Dismiss'
    }, Common.Controllers.ElixirBrowserKeys || {}));
});

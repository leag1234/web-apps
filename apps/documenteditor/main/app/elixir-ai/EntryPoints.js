/*
 * SPDX-FileCopyrightText: 2026 ElixirOffice contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 *
 * Elixir AI entry points of the Word editor, besides the command bar (Alt+I) and the Elixir tab:
 *  - AskMenu: the "Ask Elixir" menu opened from the mini toolbar sparkle and the context menus:
 *    suggestions that depend on what is selected (text, several paragraphs, a list, a table, a
 *    picture) and a free-text field;
 *  - the "/" menu: typing "/" at the start of an empty paragraph opens a quick menu (headings,
 *    lists, table, picture, page break, Ask Elixir…); Escape or any other key closes it and the
 *    "/" stays typed;
 *  - the hint "Type / for commands, or ask Elixir", shown faintly in the empty paragraph that holds
 *    the cursor (and only there).
 * The title bar search ("Search or ask Elixir") lives in common/main/lib/elixir-chrome/TitleBar.js.
 * Reads only: the selection and the current paragraph are read through the editor's own objects;
 * every change goes through the editor commands the ribbon uses, or through the command bar.
 */

define([
    'common/main/lib/elixir-ai/Core',
    'common/main/lib/elixir-ai/Icons',
    'common/main/lib/elixir-ai/Strings'
], function () {
    'use strict';

    var ElixirAI = window.Common.ElixirAI;
    ElixirAI.entryPoints = ElixirAI.entryPoints || {};
    var t = function (k) { return ElixirAI.t ? ElixirAI.t(k) : k; };

    var CSS = [
        '.elixir-pop{position:fixed;z-index:1004;min-width:260px;max-width:360px;background:var(--background-normal,#fff);color:var(--text-normal,#222);border:1px solid var(--border-regular-control,#c0c0c0);',
        '  border-radius:8px;box-shadow:0 6px 24px rgba(0,0,0,.2);font:13px/1.35 var(--font-family-base-custom,"Segoe UI",Arial,sans-serif);padding:4px 0;}',
        '.elixir-pop .elixir-pop-in{display:flex;align-items:center;gap:6px;margin:4px 8px 6px;padding:4px 6px;border:1px solid var(--border-regular-control,#c0c0c0);border-radius:6px;}',
        '.elixir-pop .elixir-pop-in input{flex:1;border:none;outline:none;background:transparent;color:inherit;font:inherit;min-width:0;}',
        '.elixir-pop .elixir-spark{width:18px;height:18px;flex:none;color:#7c4dff;}',
        '.elixir-pop .elixir-pop-it{display:flex;align-items:center;gap:8px;padding:5px 12px;cursor:pointer;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}',
        '.elixir-pop .elixir-pop-it.active,.elixir-pop .elixir-pop-it:hover{background:var(--highlight-button-hover,#e8e8e8);}',
        '.elixir-pop .elixir-pop-it .k{margin-left:auto;color:var(--text-secondary,#888);font-size:11px;}',
        '.elixir-pop .elixir-pop-sep{height:1px;margin:4px 0;background:var(--border-divider,#e0e0e0);}',
        '.elixir-slash-hint{position:fixed;z-index:3;pointer-events:none;white-space:nowrap;color:var(--text-tertiary,#9a9a9a);opacity:.75;font-family:var(--font-family-base-custom,"Segoe UI",Arial,sans-serif);display:none;}',
        '.menu__icon.elixir-ai-icon{-webkit-mask-size:20px 20px!important;mask-size:20px 20px!important;background-color:#7c4dff!important;}',
        '.mini-toolbar .elixir-ai-icon{-webkit-mask-size:20px 20px!important;mask-size:20px 20px!important;background-color:#7c4dff!important;}'
    ].join('\n');

    function injectStyles() {
        if (document.getElementById('elixir-entry-styles')) return;
        var st = document.createElement('style');
        st.id = 'elixir-entry-styles';
        st.textContent = CSS;
        document.head.appendChild(st);
    }

    function esc(s) { return String(s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); }

    // ---------------------------------------------------------------------------------------------
    // Selection probe (read only): what the Ask Elixir suggestions depend on
    // ---------------------------------------------------------------------------------------------

    function probe(api) {
        var out = { text: '', paragraphs: 0, list: false, table: false, image: false };
        try { out.text = (api.asc_GetSelectedText && api.asc_GetSelectedText()) || ''; } catch (e) {}
        try {
            var els = api.getSelectedElements ? api.getSelectedElements() : [];
            for (var i = 0; i < els.length; i++) {
                var type = els[i].get_ObjectType(), T = Asc.c_oAscTypeSelectElement;
                if (type === T.Table) out.table = true;
                if (type === T.Image) {
                    var v = els[i].get_ObjectValue();
                    if (!(v && v.get_ShapeProperties && v.get_ShapeProperties() && out.text)) out.image = true;
                }
            }
        } catch (e) {}
        try {
            var ld = api.WordControl.m_oLogicDocument;
            var ps = ld.GetSelectedParagraphs ? ld.GetSelectedParagraphs() : [];
            out.paragraphs = ps.length;
            for (var j = 0; j < ps.length; j++) if (ps[j].GetNumPr && ps[j].GetNumPr()) out.list = true;
        } catch (e) {}
        out.hasText = !!out.text.replace(/\s/g, '').length;
        return out;
    }

    // Suggestions: [{label, instruction | prefill}]
    function suggestions(p) {
        var s = [];
        if (p.image) {
            s.push({ label: t('s_inline'), instruction: t('s_inline') });
            s.push({ label: t('s_center'), instruction: t('s_center') });
            s.push({ label: t('s_wrapRight'), instruction: t('s_wrapRight') });
            return s;
        }
        if (p.table) {
            s.push({ label: t('s_headerRow'), instruction: t('s_headerRow') });
            s.push({ label: t('s_sortBy'), prefill: t('s_sortPrefill') });
        }
        if (p.hasText && p.list && !p.table) s.push({ label: t('s_toTable'), instruction: t('s_toTable') });
        if (p.hasText && p.paragraphs > 1 && !p.table) {
            s.push({ label: t('s_twoColumns'), instruction: t('s_twoColumns') });
            s.push({ label: t('s_headings'), instruction: t('s_headings') });
            if (!p.list) s.push({ label: t('s_bullets'), instruction: t('s_bullets') });
        }
        if (p.hasText) {
            if (s.length) s.push({ sep: true });
            s.push({ label: t('s_rewrite'), instruction: t('s_rewrite') });
            s.push({ label: t('s_shorten'), instruction: t('s_shorten') });
            s.push({ label: t('s_spelling'), instruction: t('s_spelling') });
            s.push({ label: t('s_translate'), prefill: t('s_translatePrefill') });
        }
        return s;
    }

    // ---------------------------------------------------------------------------------------------
    // Popup list (shared by the Ask Elixir menu and the "/" menu)
    // ---------------------------------------------------------------------------------------------

    function Popup(opts) {
        this.opts = opts;
        this.items = opts.items;
        this.active = opts.input ? -1 : 0;
        injectStyles();
        var el = this.el = document.createElement('div');
        el.className = 'elixir-pop ' + (opts.cls || '');
        el.id = opts.id;
        el.setAttribute('role', 'menu');
        var html = '';
        if (opts.input) html += '<div class="elixir-pop-in"><span class="elixir-spark">' + ElixirAI.Icons.svg('sparkle', 18) + '</span><input type="text" spellcheck="true" placeholder="' + esc(opts.placeholder || '') + '" aria-label="' + esc(opts.placeholder || '') + '"></div>';
        var n = 0;
        this.items.forEach(function (it) {
            if (it.sep) { html += '<div class="elixir-pop-sep"></div>'; return; }
            it._i = n++;
            html += '<div class="elixir-pop-it" role="menuitem" data-i="' + it._i + '"' + (it.id ? ' data-id="' + it.id + '"' : '') + '>' + esc(it.label) + (it.key ? '<span class="k">' + esc(it.key) + '</span>' : '') + '</div>';
        });
        el.innerHTML = html;
        document.body.appendChild(el);
        this.rows = Array.prototype.slice.call(el.querySelectorAll('.elixir-pop-it'));
        this.real = this.items.filter(function (it) { return !it.sep; });
        var me = this;
        el.addEventListener('mousedown', function (e) { if (e.target.tagName !== 'INPUT') e.preventDefault(); });
        this.rows.forEach(function (r) {
            r.addEventListener('click', function () { me.pick(+r.getAttribute('data-i')); });
            r.addEventListener('mousemove', function () { me.highlight(+r.getAttribute('data-i')); });
        });
        this.input = el.querySelector('input');
        this.place(opts.x, opts.y, opts.above);
        this.highlight(this.active);
    }
    Popup.prototype.place = function (x, y, above) {
        var w = this.el.offsetWidth, h = this.el.offsetHeight;
        var left = Math.max(4, Math.min(x, window.innerWidth - w - 4));
        var top = above ? y - h - 4 : y;
        if (top + h > window.innerHeight - 4) top = Math.max(4, (above ? y : y - 24) - h - 4);
        this.el.style.left = left + 'px';
        this.el.style.top = Math.max(4, top) + 'px';
    };
    Popup.prototype.highlight = function (i) {
        this.active = i;
        this.rows.forEach(function (r) { r.classList.toggle('active', +r.getAttribute('data-i') === i); });
    };
    Popup.prototype.move = function (d) {
        var n = this.real.length;
        if (!n) return;
        var i = this.active < 0 ? (d > 0 ? 0 : n - 1) : (this.active + d + n) % n;
        this.highlight(i);
    };
    Popup.prototype.pick = function (i) {
        var it = this.real[i];
        if (it) this.opts.onPick(it);
    };
    Popup.prototype.remove = function () { this.el.remove(); };

    // ---------------------------------------------------------------------------------------------
    // Setup
    // ---------------------------------------------------------------------------------------------

    function setup(ctrl) {
        injectStyles();
        var api = ctrl.api, ep = ElixirAI.entryPoints.word;
        ep.ctrl = ctrl;
        var ld = function () { try { return api.WordControl.m_oLogicDocument; } catch (e) { return null; } };
        var docFocused = function () {
            var a = document.activeElement;
            return !!a && (a.id === 'area_id' || !!$(a).closest('#editor_sdk').length);
        };
        var modal = function () {
            return (Common.Utils.ModalWindow && Common.Utils.ModalWindow.isVisible()) || $('#file-menu-panel').is(':visible');
        };
        // the empty paragraph holding the caret, or null
        var emptyCaretParagraph = function () {
            var d = ld();
            if (!d || d.IsSelectionUse && d.IsSelectionUse()) return null;
            try {
                if (d.IsHdrFtr && d.IsHdrFtr()) return null;
                var p = d.GetCurrentParagraph(false, false);
                if (!p || !p.IsEmpty || !p.IsEmpty()) return null;
                if (p.GetNumPr && p.GetNumPr()) return null;
                if (p.IsInFixedForm && p.IsInFixedForm()) return null;
                return p;
            } catch (e) { return null; }
        };
        var caretRect = function () {
            var c = document.getElementById('id_target_cursor');
            if (!c || c.style.display === 'none') return null;
            var r = c.getBoundingClientRect(), sdk = document.getElementById('editor_sdk');
            var b = sdk ? sdk.getBoundingClientRect() : null;
            if (!r.height || (b && (r.top < b.top || r.bottom > b.bottom || r.left < b.left || r.right > b.right))) return null;
            return r;
        };

        // ---- Ask Elixir menu ----
        var ask = null;
        ep.closeAskMenu = function () {
            if (!ask) return;
            ask.remove(); ask = null;
            $(document).off('mousedown.elixirask');
            api.asc_enableKeyEvents(true);
        };
        ep.openAskMenu = function (x, y, above) {
            ep.closeAskMenu();
            closeSlash(false);
            var items = suggestions(probe(api));
            ask = new Popup({
                id: 'elixir-ask-menu', cls: 'elixir-ask-menu', input: true, placeholder: t('askMenuPlaceholder'),
                items: items, x: x, y: y, above: above,
                onPick: function (it) {
                    if (it.prefill) { ask.input.value = it.prefill; ask.input.focus(); ask.highlight(-1); return; }
                    ep.closeAskMenu();
                    ctrl.ask(it.instruction);
                }
            });
            api.asc_enableKeyEvents(false);
            var input = ask.input;
            input.addEventListener('keydown', function (e) {
                e.stopPropagation();
                if (e.key === 'ArrowDown') { e.preventDefault(); ask.move(1); }
                else if (e.key === 'ArrowUp') { e.preventDefault(); ask.move(-1); }
                else if (e.key === 'Escape') { e.preventDefault(); ep.closeAskMenu(); Common.NotificationCenter.trigger('edit:complete'); }
                else if (e.key === 'Enter') {
                    e.preventDefault();
                    var v = input.value.trim(), it = ask.real[ask.active];
                    if (it && (!v || (it.prefill && v === it.prefill.trim()))) ask.pick(ask.active);
                    else if (v) { ep.closeAskMenu(); ctrl.ask(v); }
                }
            });
            input.addEventListener('input', function () { if (input.value) ask.highlight(-1); });
            ['keyup', 'keypress'].forEach(function (k) { input.addEventListener(k, function (e) { e.stopPropagation(); }); });
            setTimeout(function () {
                input && input.focus();
                $(document).on('mousedown.elixirask', function (e) { if (ask && !ask.el.contains(e.target)) ep.closeAskMenu(); });
            }, 0);
            return ask;
        };

        // ---- "/" menu ----
        var slash = null;
        var closeSlash = ep.closeSlashMenu = function () {
            if (!slash) return;
            slash.remove(); slash = null;
            updateHint();
        };
        // removes the "/" typed in the paragraph, as Backspace does (one undo step, tracked if tracking is on)
        var removeSlash = function () {
            var d = ld(), p = d && d.GetCurrentParagraph(false, false);
            if (!p || p.GetText().replace(/[\r\n]+$/, '') !== '/') return;
            var area = document.getElementById('area_id') || document.activeElement;
            var ev = new KeyboardEvent('keydown', { key: 'Backspace', code: 'Backspace', bubbles: true, cancelable: true });
            Object.defineProperty(ev, 'keyCode', { get: function () { return 8; } });
            Object.defineProperty(ev, 'which', { get: function () { return 8; } });
            area && area.dispatchEvent(ev);
            var up = new KeyboardEvent('keyup', { key: 'Backspace', code: 'Backspace', bubbles: true });
            Object.defineProperty(up, 'keyCode', { get: function () { return 8; } });
            area && area.dispatchEvent(up);
        };
        var tbc = function () { return DE.getController('Toolbar'); };
        var SLASH = [
            { id: 'h1', label: 'sl_h1', run: function () { api.put_Style('Heading 1'); } },
            { id: 'h2', label: 'sl_h2', run: function () { api.put_Style('Heading 2'); } },
            { id: 'bullets', label: 'sl_bullets', run: function () { tbc().onMarkers({ pressed: true }); } },
            { id: 'numbers', label: 'sl_numbers', run: function () { tbc().onNumbers({ pressed: true }); } },
            { id: 'table', label: 'sl_table', run: function () { tbc().onInsertTableClick(null, { value: 'custom' }); } },
            {
                id: 'picture', label: 'sl_picture', run: function () {
                    var tb = tbc().toolbar, b = tb.btnsInsertImage && tb.btnsInsertImage[0];
                    var it = b && b.menu && _.find(b.menu.items, function (i) { return i.value === 'file'; });
                    if (it) b.menu.trigger('item:click', b.menu, it, {});
                }
            },
            { id: 'pagebreak', label: 'sl_pageBreak', run: function () { api.put_AddPageBreak(); } },
            { id: 'ask', label: 'sl_ask', sepBefore: true, run: function () { ctrl.openBar(); } }
        ];
        var openSlash = function () {
            var r = caretRect();
            if (!r) return;
            hideHint();
            var items = [];
            SLASH.forEach(function (s) { if (s.sepBefore) items.push({ sep: true }); items.push({ id: s.id, label: t(s.label), run: s.run }); });
            slash = new Popup({
                id: 'elixir-slash-menu', cls: 'elixir-slash-menu', items: items, x: r.left, y: r.bottom + 4,
                onPick: function (it) {
                    closeSlash();
                    removeSlash();
                    setTimeout(function () { it.run(); if (it.id !== 'ask' && it.id !== 'table' && it.id !== 'picture') Common.NotificationCenter.trigger('edit:complete'); }, 30);
                }
            });
        };
        // capture phase: before the editor's own key handling
        window.addEventListener('keydown', function (e) {
            if (slash) {
                if (e.key === 'ArrowDown' || e.key === 'ArrowUp') { e.preventDefault(); e.stopImmediatePropagation(); slash.move(e.key === 'ArrowDown' ? 1 : -1); return; }
                if (e.key === 'Enter' || e.key === 'Tab') { e.preventDefault(); e.stopImmediatePropagation(); slash.pick(slash.active); return; }
                if (e.key === 'Escape') { e.preventDefault(); e.stopImmediatePropagation(); closeSlash(); return; }
                if (['Shift', 'Control', 'Alt', 'Meta', 'CapsLock', 'AltGraph'].indexOf(e.key) >= 0) return;
                closeSlash();   // any other key: typed normally, after the "/"
                return;
            }
            if (e.key !== '/' || e.ctrlKey || e.metaKey || (e.altKey && !e.getModifierState('AltGraph')) || e.isComposing) return;
            if (!ctrl.isActive() || !docFocused() || modal() || !emptyCaretParagraph()) return;
            // the "/" is typed by the editor; the menu opens once it is in the paragraph
            setTimeout(function () {
                var d = ld(), p = d && d.GetCurrentParagraph(false, false);
                if (p && p.GetText().replace(/[\r\n]+$/, '') === '/' && !(d.IsSelectionUse && d.IsSelectionUse())) openSlash();
            }, 40);
        }, true);
        $(document).on('mousedown', function (e) { if (slash && !slash.el.contains(e.target)) closeSlash(); });

        // ---- hint in the empty paragraph holding the caret ----
        var hint = document.createElement('div');
        hint.className = 'elixir-slash-hint';
        hint.id = 'elixir-slash-hint';
        hint.textContent = t('slashHint');
        document.body.appendChild(hint);
        var hideHint = function () { hint.style.display = 'none'; };
        var timer = null;
        var updateHint = ep.updateHint = function () {
            clearTimeout(timer);
            timer = setTimeout(function () {
                if (slash || ask || !ctrl.isActive() || modal() || !docFocused() || !emptyCaretParagraph()) return hideHint();
                var r = caretRect();
                if (!r) return hideHint();
                hint.style.fontSize = Math.max(11, Math.min(20, Math.round(r.height * 0.72))) + 'px';
                hint.style.left = Math.round(r.right + 3) + 'px';
                hint.style.top = Math.round(r.top + (r.height - hint.offsetHeight || 0) / 2) + 'px';
                hint.style.display = 'block';
                var h = hint.offsetHeight;
                hint.style.top = Math.round(r.top + (r.height - h) / 2) + 'px';
            }, 60);
        };
        api.asc_registerCallback('asc_onFocusObject', updateHint);
        api.asc_registerCallback('asc_onDocumentContentReady', updateHint);
        window.addEventListener('keyup', updateHint, true);
        window.addEventListener('mouseup', updateHint, true);
        window.addEventListener('focusin', updateHint, true);
        window.addEventListener('focusout', updateHint, true);
        $('#editor_sdk').on('wheel', function () { hideHint(); updateHint(); });
        Common.NotificationCenter.on('window:resize', updateHint);
        updateHint();
    }

    ElixirAI.entryPoints.word = { setup: setup, probe: probe, suggestions: suggestions };
    return ElixirAI.entryPoints.word;
});

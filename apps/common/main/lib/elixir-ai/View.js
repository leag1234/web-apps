/*
 * SPDX-FileCopyrightText: 2026 ElixirOffice contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 *
 * Elixir AI user interface: the floating command bar over the document and the settings /
 * journal dialog. Plain DOM, self-contained styles (injected once), theme colors through the
 * editor's CSS variables, so the module does not touch the editor's LESS or templates.
 */

define([
    'common/main/lib/elixir-ai/Core',
    'common/main/lib/elixir-ai/Icons',
    'common/main/lib/elixir-ai/Strings'
], function () {
    'use strict';

    var ElixirAI = window.Common.ElixirAI;

    // ---------------------------------------------------------------------------------------------
    // Strings (en, fr)
    // ---------------------------------------------------------------------------------------------

    var STRINGS = {
        en: {
            tab: 'Elixir', ask: 'Ask Elixir', settings: 'AI settings', journal: 'AI journal',
            askHint: 'Ask Elixir to change the document structure', settingsHint: 'Model connection and AI journal',
            placeholder: 'Ask Elixir to change this document… e.g. “Put this section in two columns”',
            scopeSelection: 'Selection', scopeDocument: 'Whole document',
            send: 'Send', stop: 'Stop', accept: 'Accept', discard: 'Discard', close: 'Close',
            notConfigured: 'Elixir AI is not configured: no request is sent until you set a model endpoint.',
            configure: 'Configure…', working: 'Working…', reading: 'Reading the document…',
            done: 'Done. Changes are shown as tracked changes.', doneUntracked: 'Done.', noChange: 'No change was made.',
            accepted: 'Changes accepted.', discarded: 'Changes discarded.',
            discardedReject: 'The document changed since the AI action: tracked changes were rejected, untracked layout changes were kept.',
            failed: 'The action failed and was reverted: ', cancelled: 'Cancelled. Nothing was changed.',
            untracked: 'Not trackable in Word',
            shortcut: 'Shortcut', footer: 'Context sent: the selection, or the whole document when nothing is selected.',
            // settings dialog
            dlgTitle: 'Elixir AI settings', endpoint: 'Endpoint URL (OpenAI-compatible)', apiKey: 'API key', model: 'Model',
            extra: 'Extra request parameters (JSON, optional)', rounds: 'Maximum tool rounds',
            privacy: 'Stored in this browser only (localStorage). The key is sent only to the endpoint above, never elsewhere, never logged.',
            save: 'Save', test: 'Test connection', clear: 'Remove settings', saved: 'Saved.',
            testing: 'Testing…', testOk: 'Connection OK', journalTitle: 'AI journal (stored locally in this browser)',
            exportJ: 'Export JSON', clearJ: 'Clear journal', emptyJ: 'No AI call yet.', invalidUrl: 'The endpoint must start with http:// or https://',
            invalidJson: 'Extra parameters must be a JSON object.',
            when: 'When', prompt: 'Instruction', outcome: 'Result', time: 'Time', tokens: 'Tokens', cost: 'Est. cost'
        },
        fr: {
            tab: 'Elixir', ask: 'Demander à Elixir', settings: 'Réglages IA', journal: 'Journal IA',
            askHint: 'Demander à Elixir de modifier la structure du document', settingsHint: 'Connexion au modèle et journal IA',
            placeholder: 'Demandez à Elixir de modifier ce document… ex. « Mets cette section en deux colonnes »',
            scopeSelection: 'Sélection', scopeDocument: 'Document entier',
            send: 'Envoyer', stop: 'Arrêter', accept: 'Accepter', discard: 'Annuler les changements', close: 'Fermer',
            notConfigured: 'Elixir IA n’est pas configurée : aucune requête n’est envoyée tant qu’aucun modèle n’est réglé.',
            configure: 'Configurer…', working: 'En cours…', reading: 'Lecture du document…',
            done: 'Terminé. Les changements apparaissent en suivi des modifications.', doneUntracked: 'Terminé.', noChange: 'Aucun changement.',
            accepted: 'Changements acceptés.', discarded: 'Changements annulés.',
            discardedReject: 'Le document a changé depuis l’action IA : les révisions ont été refusées, la mise en page non suivie a été conservée.',
            failed: 'L’action a échoué et a été annulée : ', cancelled: 'Interrompu. Rien n’a été modifié.',
            untracked: 'Non suivable dans Word',
            shortcut: 'Raccourci', footer: 'Contexte envoyé : la sélection, ou tout le document si rien n’est sélectionné.',
            dlgTitle: 'Réglages Elixir IA', endpoint: 'URL du service (compatible OpenAI)', apiKey: 'Clé d’API', model: 'Modèle',
            extra: 'Paramètres de requête supplémentaires (JSON, facultatif)', rounds: 'Nombre maximal de tours d’outils',
            privacy: 'Enregistré dans ce navigateur uniquement (localStorage). La clé n’est envoyée qu’au service ci-dessus, jamais ailleurs, jamais journalisée.',
            save: 'Enregistrer', test: 'Tester la connexion', clear: 'Effacer les réglages', saved: 'Enregistré.',
            testing: 'Test en cours…', testOk: 'Connexion OK', journalTitle: 'Journal IA (stocké localement dans ce navigateur)',
            exportJ: 'Exporter (JSON)', clearJ: 'Vider le journal', emptyJ: 'Aucun appel IA pour l’instant.', invalidUrl: 'L’URL doit commencer par http:// ou https://',
            invalidJson: 'Les paramètres supplémentaires doivent être un objet JSON.',
            when: 'Date', prompt: 'Consigne', outcome: 'Résultat', time: 'Durée', tokens: 'Jetons', cost: 'Coût estimé'
        }
    };

    var lang = 'en', requestedLang = 'en';
    ElixirAI.setLang = function (l) { requestedLang = l; lang = ElixirAI.langOf(l); };
    ElixirAI.getLang = function () { return lang; };
    // Command bar and entry points: Strings.js (12 languages); settings dialog: STRINGS above (en, fr)
    var t = ElixirAI.t = function (k) {
        if (k === 'invalidUrl') return ElixirAI.endpointPolicyMessage(requestedLang);
        var S = ElixirAI.Strings || {};
        return (S[lang] && S[lang][k]) || (STRINGS[lang] && STRINGS[lang][k]) || (S.en && S.en[k]) || STRINGS.en[k] || k;
    };
    // t() with {0} / {name} placeholders
    var tf = ElixirAI.tf = function (k, vals) {
        vals = vals || {};
        return t(k).replace(/\{(\w+)\}/g, function (m, n) { return vals[n] !== undefined && vals[n] !== null ? String(vals[n]) : m; });
    };

    function esc(s) {
        return String(s === undefined || s === null ? '' : s).replace(/[&<>"']/g, function (c) {
            return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
        });
    }

    // ---------------------------------------------------------------------------------------------
    // Styles
    // ---------------------------------------------------------------------------------------------

    var CSS = [
        '.elixir-ai-cfg .lead{margin:0 0 10px;color:var(--text-secondary,#555)}.elixir-ai-cfg select{width:100%;box-sizing:border-box;padding:6px 8px}.elixir-ai-cfg .elixir-ai-managed{margin:0 0 10px;padding:6px 10px;border-radius:6px;background:var(--highlight-button-hover,#eef3fb)}',
        '.elixir-ai-notice{margin:10px 0 4px;padding:8px 10px;border-left:3px solid var(--border-preview-select,#4a7fd1);background:var(--background-toolbar,#f4f4f4);font-weight:600}.elixir-ai-adv{margin-top:8px}.elixir-ai-adv summary{cursor:pointer;color:var(--text-secondary,#555)}',
        '.elx-page .elixir-ai-cfg{max-width:560px}.elx-page .elixir-ai-cfg label{display:block;margin:10px 0 3px}.elx-page .elixir-ai-cfg input,.elx-page .elixir-ai-cfg select{width:100%;box-sizing:border-box;padding:6px 8px}',
        '.elixir-ai-icon{background:none!important;background-color:var(--text-normal,#333)!important;-webkit-mask:var(--elixir-mask) center/24px 24px no-repeat;mask:var(--elixir-mask) center/24px 24px no-repeat;}',
        '.elixir-ai-bar{position:fixed;z-index:1001;width:640px;max-width:calc(100vw - 32px);box-sizing:border-box;background:var(--background-normal,#fff);color:var(--text-normal,#222);',
        '  border:1px solid var(--border-regular-control,#c0c0c0);border-radius:12px;box-shadow:0 8px 28px rgba(0,0,0,.22);font:13px/1.4 var(--font-family-base-custom,"Segoe UI",Arial,sans-serif);padding:10px 12px;}',
        '.elixir-ai-bar *{box-sizing:border-box;}',
        '.elixir-ai-row{display:flex;align-items:flex-start;gap:8px;}',
        '.elixir-ai-spark{flex:none;width:24px;height:24px;margin-top:4px;color:#7c4dff;}',
        '.elixir-ai-bar.busy .elixir-ai-spark{animation:elixir-ai-pulse 1.2s ease-in-out infinite;}',
        '@keyframes elixir-ai-pulse{0%,100%{opacity:1}50%{opacity:.35}}',
        '.elixir-ai-input{flex:1;resize:none;border:none;outline:none;background:transparent;color:inherit;font:inherit;font-size:14px;min-height:32px;max-height:120px;padding:5px 2px;}',
        '.elixir-ai-btn{flex:none;border:1px solid var(--border-regular-control,#c0c0c0);background:var(--background-normal,#fff);color:inherit;border-radius:6px;padding:4px 12px;font:inherit;cursor:pointer;height:28px;}',
        '.elixir-ai-btn:hover{background:var(--highlight-button-hover,#e0e0e0);}',
        '.elixir-ai-btn.primary{background:var(--background-accent-button,#446995);border-color:var(--background-accent-button,#446995);color:var(--text-inverse,#fff);}',
        '.elixir-ai-btn[disabled]{opacity:.5;cursor:default;}',
        '.elixir-ai-meta{display:flex;justify-content:space-between;gap:8px;margin-top:4px;color:var(--text-secondary,#777);font-size:11px;}',
        '.elixir-ai-chip{border:1px solid var(--border-divider,#ddd);border-radius:10px;padding:0 8px;}',
        '.elixir-ai-panel{margin-top:8px;border-top:1px solid var(--border-divider,#e0e0e0);padding-top:8px;display:none;}',
        '.elixir-ai-panel.show{display:block;}',
        '.elixir-ai-plan{color:var(--text-secondary,#666);white-space:pre-wrap;margin-bottom:4px;}',
        '.elixir-ai-steps{list-style:none;margin:0;padding:0;max-height:180px;overflow:auto;}',
        '.elixir-ai-steps li{padding:2px 0 2px 20px;position:relative;}',
        '.elixir-ai-steps li:before{content:"";position:absolute;left:4px;top:8px;width:8px;height:8px;border-radius:50%;background:var(--border-regular-control,#bbb);}',
        '.elixir-ai-steps li.ok:before{background:#2e7d32;}.elixir-ai-steps li.err:before{background:#c62828;}',
        '.elixir-ai-steps li .err-msg{color:#c62828;font-size:11px;display:block;}',
        '.elixir-ai-status{margin-top:6px;}',
        '.elixir-ai-status.error{color:#c62828;}',
        '.elixir-ai-answer{margin-top:6px;white-space:pre-wrap;}',
        '.elixir-ai-note.warn{background:#fff4e5;color:#7a4b00;}',
        '.elixir-ai-note{margin-top:6px;padding:4px 8px;border-radius:6px;background:var(--background-toolbar-additional,#f3f3f3);font-size:12px;}',
        '.elixir-ai-actions{display:flex;justify-content:flex-end;gap:8px;margin-top:8px;}',
        '.elixir-ai-modal-mask{position:fixed;inset:0;background:rgba(0,0,0,.25);z-index:1002;}',
        '.elixir-ai-dlg{position:fixed;z-index:1003;left:50%;top:50%;transform:translate(-50%,-50%);width:760px;max-width:calc(100vw - 32px);max-height:calc(100vh - 48px);overflow:auto;',
        '  background:var(--background-normal,#fff);color:var(--text-normal,#222);border:1px solid var(--border-regular-control,#c0c0c0);border-radius:8px;box-shadow:0 10px 40px rgba(0,0,0,.3);font:13px/1.4 var(--font-family-base-custom,"Segoe UI",Arial,sans-serif);padding:16px 20px;}',
        '.elixir-ai-dlg h2{font-size:15px;margin:0 0 12px;font-weight:600;} .elixir-ai-dlg h3{font-size:13px;margin:18px 0 8px;font-weight:600;}',
        '.elixir-ai-dlg label{display:block;margin:8px 0 3px;color:var(--text-secondary,#555);}',
        '.elixir-ai-dlg input,.elixir-ai-dlg textarea{width:100%;box-sizing:border-box;border:1px solid var(--border-regular-control,#c0c0c0);border-radius:4px;padding:5px 7px;font:inherit;background:var(--background-normal,#fff);color:inherit;}',
        '.elixir-ai-dlg .hint{font-size:11px;color:var(--text-secondary,#777);margin-top:6px;}',
        '.elixir-ai-dlg .msg{margin-left:8px;font-size:12px;} .elixir-ai-dlg .msg.error{color:#c62828;}',
        '.elixir-ai-dlg table{width:100%;border-collapse:collapse;font-size:12px;} .elixir-ai-dlg td,.elixir-ai-dlg th{border-bottom:1px solid var(--border-divider,#e5e5e5);padding:4px 6px;text-align:left;vertical-align:top;}',
        '.elixir-ai-dlg pre{white-space:pre-wrap;word-break:break-word;font-size:11px;max-height:240px;overflow:auto;background:var(--background-toolbar-additional,#f5f5f5);padding:6px;margin:4px 0;}',
        '.elixir-ai-dlg tr.entry{cursor:pointer;} .elixir-ai-dlg tr.entry:hover{background:var(--highlight-button-hover,#eee);}'
    ].join('\n');

    function injectStyles() {
        if (document.getElementById('elixir-ai-styles')) return;
        var st = document.createElement('style');
        st.id = 'elixir-ai-styles';
        st.textContent = CSS + '\n' + ['sparkle', 'settings', 'history'].map(function (n) {
            return '.elixir-ai-icon-' + n + '{--elixir-mask:' + ElixirAI.Icons.maskUrl(n) + ';}';
        }).join('\n');
        document.head.appendChild(st);
    }
    ElixirAI.injectStyles = injectStyles;

    // ---------------------------------------------------------------------------------------------
    // Command bar
    // ---------------------------------------------------------------------------------------------

    // Label of a step while it runs (reading tools) and plain words of what it did (changes)
    var READ_LABELS = {
        read_context: 'r_read', document_info: 'r_info', find_text: 'r_find', explain_layout: 'r_layout',
        get_outline_candidates: 'r_outline', get_layout_summary: 'r_layoutSummary', find_overflowing_elements: 'r_overflow'
    };
    function runLabel(name, args) {
        if (ElixirAI.slideText && ['read_slides','summarize_deck'].indexOf(name)>=0) return ElixirAI.slideText(name==='read_slides'?'read':'summarize');
        if (['read_range','explain_formula','summarize_range'].indexOf(name)>=0 && ElixirAI.cellText) return ElixirAI.cellText(name==='read_range'?'read':name==='explain_formula'?'explain':'summarize');
        if (READ_LABELS[name]) return tf(READ_LABELS[name], { find: (args && (args.text || args.find)) || '' });
        if (name === 'submit_plan') return tf('r_plan', [args && args.goal || '']).replace('{0}', args && args.goal || '');
        return t('working');
    }
    function doneText(name, args, result) {
        var d = (result && result.done) || {}, a = args || {};
        var n = d.n;
        var slideKeys={create_slides:'create',rewrite_text:'rewrite',set_layout:'layout',apply_theme:'theme',align_shapes:'align',set_notes:'notes'};
        if (ElixirAI.slideText && slideKeys[name]) return ElixirAI.slideText(slideKeys[name],{n:n});
        var cellKeys={write_formula:'formula',sort_range:'sort',filter_range:'filter',create_chart:'chart',create_pivot:'pivot'};
        if ((cellKeys[name] || name==='clean_range') && ElixirAI.cellText) return ElixirAI.cellText(cellKeys[name] || a.operation,{range:d.range || a.range,n:n});
        switch (name) {
            case 'apply_styles':
                return Object.keys(d.byStyle || {}).map(function (st) { return tf('d_apply_style', { style: st, n: d.byStyle[st] }); }).join('; ');
            case 'apply_style': return tf('d_apply_style', { style: a.style, n: n });
            case 'set_list': return tf(d.kind === 'bullet' ? 'd_list_bullet' : d.kind === 'none' ? 'd_list_none' : 'd_list_numbered', { n: n });
            case 'insert_table_of_contents': return tf('d_toc', { n: n });
            case 'table_rows': return tf(d.action === 'delete' ? 'd_rows_delete' : 'd_rows_insert', { n: n });
            case 'table_columns': return tf(d.action === 'delete' ? 'd_cols_delete' : 'd_cols_insert', { n: n });
            case 'table_merge_cells': return tf('d_merge', { n: n });
            case 'table_sort': return tf('d_sort', { column: d.column });
            case 'table_set_cells': return tf('d_cells', { n: n });
            case 'set_image_layout': return t('d_image_layout');
            case 'fit_content': return tf('d_fit', { n: n });
            case 'set_page_setup': return t('d_page_setup');
            case 'set_columns': return tf('d_columns', { n: n, count: d.count });
            case 'insert_break': return t(d.kind === 'page' ? 'd_break_page' : d.kind === 'column' ? 'd_break_column' : 'd_break_section');
            case 'set_header_footer': return t(d.part === 'header' ? 'd_header' : 'd_footer');
            case 'add_comment': return t('d_comment');
            case 'list_to_table': case 'create_table': return tf('d_' + name, { rows: d.rows, cols: d.cols });
            case 'find_replace': return tf('d_find_replace', { find: d.find, replace: d.replace, n: n });
            case 'table_format': return t('d_table_format');
            case 'insert_image': return t('d_insert_image');
            case 'modify_style': return tf('d_modify_style', { style: a.style });
        }
        var k = 'd_' + name;
        return t(k) !== k ? tf(k, { n: n }) : name;
    }
    ElixirAI.doneText = doneText;
    // The plain-words summary of an action: one line per kind of change, in order
    ElixirAI.summarize = function (done) {
        var out = [];
        (done || []).forEach(function (x) { var s = doneText(x.name, x.args, x.result); if (s && out.indexOf(s) < 0) out.push(s); });
        return out.join('. ') + (out.length ? '.' : '');
    };

    /**
     * host: {api, toolsetFactory(), openSettings(), shortcutLabel, isHistoryAt(index)…}
     */
    function setGuidedTipsSuspended(owner, suspended) {
        var manager = Common.UI && Common.UI.TooltipManager;
        var action = suspended ? 'suspend' : 'resume';
        if (manager && typeof manager[action] === 'function') manager[action](owner);
    }

    function CommandBar(host) {
        this.host = host;
        this.el = null;
        this.setState('closed');
        this.result = null;
    }

    CommandBar.prototype.render = function () {
        if (this.el) return;
        injectStyles();
        var el = this.el = document.createElement('div');
        el.className = 'elixir-ai-bar';
        el.id = 'elixir-ai-bar';
        el.setAttribute('role', 'dialog');
        el.setAttribute('aria-label', 'Elixir AI');
        el.innerHTML =
            '<div class="elixir-ai-row">' +
                '<span class="elixir-ai-spark">' + ElixirAI.Icons.svg('sparkle', 24) + '</span>' +
                '<textarea class="elixir-ai-input" rows="1" spellcheck="true"></textarea>' +
                '<button class="elixir-ai-btn primary" data-act="send"></button>' +
            '</div>' +
            '<div class="elixir-ai-meta"><span class="elixir-ai-chip elixir-ai-scope"></span><span class="elixir-ai-footer"></span></div>' +
            '<div class="elixir-ai-panel">' +
                '<div class="elixir-ai-plan"></div>' +
                '<ul class="elixir-ai-steps"></ul>' +
                '<div class="elixir-ai-status"></div>' +
                '<div class="elixir-ai-notes"></div>' +
                '<div class="elixir-ai-actions"></div>' +
            '</div>';
        document.body.appendChild(el);
        this.input = el.querySelector('.elixir-ai-input');
        this.btnSend = el.querySelector('[data-act=send]');
        this.panel = el.querySelector('.elixir-ai-panel');

        var me = this;
        this.input.placeholder = this.host.toolsetId==='slide' ? ElixirAI.slideText('placeholder') : this.host.toolsetId==='cell' ? ElixirAI.cellText('placeholder') : t('placeholder');
        this.btnSend.textContent = t('send');
        el.querySelector('.elixir-ai-footer').textContent = this.host.toolsetId==='slide' ? ElixirAI.slideText('footer') : this.host.toolsetId==='cell' ? ElixirAI.cellText('footer') : t('footer');
        this.btnSend.addEventListener('click', function () { me.state === 'running' ? me.stop() : me.submit(); });
        this.input.addEventListener('input', function () {
            me.input.style.height = 'auto';
            me.input.style.height = Math.min(120, me.input.scrollHeight) + 'px';
        });
        // Keys typed in the bar belong to the bar: they must not reach the editor shortcuts.
        el.addEventListener('keydown', function (e) {
            if (e.key === 'Escape') { e.preventDefault(); me.state === 'running' ? me.stop() : me.close(); }
            else if (e.key === 'Enter' && !e.shiftKey && e.target === me.input) { e.preventDefault(); if (me.state !== 'running') me.submit(); }
            e.stopPropagation();
        });
        el.addEventListener('keyup', function (e) { e.stopPropagation(); });
        el.addEventListener('keypress', function (e) { e.stopPropagation(); });
        this.input.addEventListener('focus', function () { me.host.api && me.host.api.asc_enableKeyEvents(false); });
        // Escape closes the bar even when the focus went back to the document
        document.addEventListener('keydown', function (e) {
            if (e.key !== 'Escape' || !me.isOpen() || me.el.contains(e.target)) return;
            if (Common.Utils && Common.Utils.ModalWindow && Common.Utils.ModalWindow.isVisible()) return;
            me.state === 'running' ? me.stop() : me.close();
        }, true);
        window.addEventListener('resize', function () { me.position(); });
    };

    CommandBar.prototype.position = function () {
        if (!this.el) return;
        var sdk = document.getElementById('editor_sdk');
        var r = sdk ? sdk.getBoundingClientRect() : { left: 0, top: 80, width: window.innerWidth, bottom: window.innerHeight };
        var w = Math.min(640, window.innerWidth - 32);
        // Docked at the bottom of the document area, so the selection it works on stays visible;
        // the progress panel grows upwards.
        this.el.style.left = Math.max(16, r.left + (r.width - w) / 2) + 'px';
        this.el.style.top = 'auto';
        this.el.style.bottom = Math.max(16, window.innerHeight - r.bottom + 28) + 'px';
    };

    CommandBar.prototype.isOpen = function () { return this.state !== 'closed'; };

    // The state is mirrored in data-state for automation and styling
    CommandBar.prototype.setState = function (state) {
        this.state = state;
        if (this.el) this.el.setAttribute('data-state', state);
    };

    CommandBar.prototype.open = function () {
        this.render();
        this.position();
        this.el.style.display = '';
        setGuidedTipsSuspended(this, true);
        if (this.state === 'closed') this.reset();
        var configured = ElixirAI.Settings.isConfigured();
        if (!configured && !this.managedChecked) {
            // the server-managed config may still be loading: reopen once it is known, unless the
            // bar was closed in the meantime (the keys typed since must not land in the bar)
            this.managedChecked = true;
            var me0 = this;
            ElixirAI.Settings.loadManaged().then(function () { if (me0.isOpen()) me0.open(); });
            return;
        }
        var sel = '';
        try { sel = this.host.api.asc_GetSelectedText ? this.host.api.asc_GetSelectedText() : ''; } catch (e) {}
        this.el.querySelector('.elixir-ai-scope').textContent = this.host.toolsetId==='slide' ? ElixirAI.slideText('scope') : this.host.toolsetId==='cell' ? ElixirAI.cellText('scope') : (sel && sel.replace(/\s/g, '').length) ? t('scopeSelection') : t('scopeDocument');
        if (!configured) {
            this.setState('unconfigured');
            this.showPanel();
            this.setStatus(t('notConfigured'));
            this.setActions([{ label: t('configure'), primary: true, fn: this.host.openSettings }, { label: t('close'), fn: this.close.bind(this) }]);
            this.input.disabled = true;
            this.btnSend.disabled = true;
        } else {
            if (this.state === 'unconfigured') this.reset();
            if (this.state !== 'review' && this.state !== 'running') this.setState('idle');
            this.input.disabled = this.state === 'running';
            this.btnSend.disabled = false;
        }
        var me = this;
        setTimeout(function () {
            var first = me.el.querySelector('.elixir-ai-actions .elixir-ai-btn');
            if (!me.input.disabled) me.input.focus();
            else if (first) first.focus();
        }, 0);
    };

    CommandBar.prototype.reset = function () {
        this.setState('idle');
        this.result = null;
        this.input.value = '';
        this.input.disabled = false;
        this.input.style.height = '';
        this.btnSend.textContent = t('send');
        this.btnSend.disabled = false;
        this.el.classList.remove('busy');
        this.panel.classList.remove('show');
        this.panel.querySelector('.elixir-ai-plan').textContent = '';
        this.panel.querySelector('.elixir-ai-steps').innerHTML = '';
        this.panel.querySelector('.elixir-ai-notes').innerHTML = '';
        this.setStatus('');
        this.setActions([]);
    };

    CommandBar.prototype.close = function () {
        if (!this.el || this.state === 'closed') return;
        if (this.state === 'running') return;
        // Closing while reviewing keeps the AI changes as ordinary tracked changes in the document.
        this.el.style.display = 'none';
        this.setState('closed');
        setGuidedTipsSuspended(this, false);
        this.result = null;
        if (this.host.api) this.host.api.asc_enableKeyEvents(true);
        this.host.onClosed && this.host.onClosed();
    };

    CommandBar.prototype.showPanel = function () { this.panel.classList.add('show'); };

    CommandBar.prototype.setStatus = function (text, isError) {
        var s = this.panel.querySelector('.elixir-ai-status');
        s.textContent = text || '';
        s.classList.toggle('error', !!isError);
    };

    CommandBar.prototype.setActions = function (list) {
        var box = this.panel.querySelector('.elixir-ai-actions');
        box.innerHTML = '';
        list.forEach(function (a) {
            var b = document.createElement('button');
            b.className = 'elixir-ai-btn' + (a.primary ? ' primary' : '');
            b.textContent = a.label;
            if (a.id) b.setAttribute('data-act', a.id);
            b.addEventListener('click', a.fn);
            box.appendChild(b);
        });
    };

    CommandBar.prototype.addStep = function (tc) {
        if (tc.name === 'cannot_do' || tc.name === 'ask_user' || tc.name === 'verify_result') return;
        if (tc.name === 'submit_plan') {
            this.panel.querySelector('.elixir-ai-plan').textContent = (tc.args && tc.args.goal) || '';
            return;
        }
        var li = document.createElement('li');
        li.textContent = runLabel(tc.name, tc.args);
        this.panel.querySelector('.elixir-ai-steps').appendChild(li);
        tc._li = li;
    };

    CommandBar.prototype.endStep = function (tc) {
        var li = tc._li;
        if (!li) return;
        var ok = tc.result && tc.result.ok !== false;
        li.className = ok ? 'ok' : 'err';
        if (ok && !READ_LABELS[tc.name]) li.textContent = doneText(tc.name, tc.args, tc.result);
        if (!ok) {
            var m = document.createElement('span');
            m.className = 'err-msg';
            m.textContent = (tc.result && tc.result.error && tc.result.error.message) || 'error';
            li.appendChild(m);
        }
        delete tc._li;
    };

    CommandBar.prototype.submit = function () {
        var text = this.input.value.trim();
        if (!text || this.state === 'running') return;
        if (this.state === 'review') this.finishReview(); // a new request keeps the previous result as tracked changes
        var me = this, s = ElixirAI.Settings.get();
        if (!ElixirAI.Settings.isConfigured(s)) { this.open(); return; }
        var toolset = this.host.toolsetFactory();
        this.reset();
        this.input.value = text;
        this.setState('running');
        this.input.disabled = true;
        this.btnSend.textContent = t('stop');
        this.el.classList.add('busy');
        this.showPanel();
        this.setStatus(t('reading'));
        this.abort = new AbortController();
        this.lastInstruction = text;
        ElixirAI.Orchestrator.run({
            instruction: text, toolset: toolset, settings: s, signal: this.abort.signal, lang: lang,
            onEvent: function (type, data) {
                if (type === 'context') {
                    me.el.querySelector('.elixir-ai-scope').textContent = me.host.toolsetId==='slide' ? ElixirAI.slideText(data.scope==='deck'?'scopeDeck':'scopeSlides',{count:data.scopeSlides.length}) : data.scope === 'range' && ElixirAI.cellText ? ElixirAI.cellText('scopeRange',data) : data.scope === 'selection' ? t('scopeSelection') : t('scopeDocument');
                    me.setStatus(t('working'));
                } else if (type === 'delta') {
                    me.panel.querySelector('.elixir-ai-plan').textContent = data.text;
                } else if (type === 'verify') me.setStatus(t('verifying'));
                else if (type === 'tool:start') me.addStep(data);
                else if (type === 'tool:end') me.endStep(data);
            }
        }).then(function (res) {
            me.abort = null;
            me.el.classList.remove('busy');
            me.btnSend.textContent = t('send');
            me.input.disabled = false;
            if (!res.ok) {
                me.setState('idle');
                var aborted = res.entry && res.entry.outcome === 'cancelled';
                me.setStatus(res.rollbackUnverified && me.host.toolsetId==='slide' ? ElixirAI.slideText('rollbackUnverified') : res.rollbackUnverified && ElixirAI.cellText ? ElixirAI.cellText('rollbackUnverified') : aborted ? t('cancelled') : t('failed') + (res.error && res.error.message || res.error), !!res.rollbackUnverified || !aborted);
                me.setActions([{ label: t('close'), fn: me.close.bind(me) }]);
                return;
            }
            me.result = res;
            var planEl = me.panel.querySelector('.elixir-ai-plan');
            var notes = me.panel.querySelector('.elixir-ai-notes');
            notes.innerHTML = '';
            res.untracked.forEach(function (n) {
                var d = document.createElement('div');
                d.className = 'elixir-ai-note';
                d.textContent = t('untracked') + ' — ' + t('u_' + n);
                notes.appendChild(d);
            });
            var again = function (instruction) { return function () { me.input.value = instruction; me.submit(); }; };
            if (res.cannotDo && !res.changed) {
                // never a silent no-op: what cannot be done, and the closest action in one click
                me.setState('idle');
                planEl.textContent = res.cannotDo.message || '';
                me.el.querySelector('.elixir-ai-status').setAttribute('data-kind', 'cannot-do');
                me.setStatus(t('cannotDo'));
                var acts = [{ label: t('close'), fn: me.close.bind(me) }];
                if (res.cannotDo.suggestion) acts.push({ id: 'suggestion', label: tf('tryThis', [res.cannotDo.suggestion]).replace('{0}', res.cannotDo.suggestion), primary: true, fn: again(res.cannotDo.suggestion) });
                me.setActions(acts);
                return;
            }
            if (res.question && !res.changed) {
                me.setState('idle');
                planEl.textContent = res.question.question || '';
                me.setStatus(t('question'));
                me.setActions((res.question.options || []).map(function (o, i) {
                    return { id: 'option-' + i, label: o, primary: i === 0, fn: again(text + ' — ' + o) };
                }).concat([{ label: t('close'), fn: me.close.bind(me) }]));
                return;
            }
            var summary = ElixirAI.summarize(res.done);
            me.summary = summary;
            // steps the model corrected itself (e.g. a refused argument) are noise once the action succeeded
            if (res.changed || res.cannotDo) Array.prototype.forEach.call(me.panel.querySelectorAll('.elixir-ai-steps li.err'), function (li) { li.remove(); });
            if (!res.changed) {
                me.setState('idle');
                // a question about the document: the model's answer; otherwise say nothing was changed
                planEl.textContent = res.answer || '';
                me.setStatus(res.answer ? '' : t('noChange'));
                me.setActions([{ label: t('close'), fn: me.close.bind(me) }]);
                return;
            }
            // the plan line keeps what Elixir understood (the plan goal); the steps say, in plain words, what was done
            if (!res.plan || !planEl.textContent) planEl.textContent = summary;
            me.el.setAttribute('data-summary', summary);
            (res.userNotes || []).forEach(function (k) {
                var un = document.createElement('div');
                un.className = 'elixir-ai-answer elixir-ai-usernote';
                un.textContent = t(k);
                notes.appendChild(un);
            });
            if (res.answer) {
                var ans = document.createElement('div');
                ans.className = 'elixir-ai-answer';
                ans.textContent = res.answer;
                notes.appendChild(ans);
            }
            if (res.verification && res.verification.goal_met === false && (res.verification.problems || []).length) {
                var w = document.createElement('div');
                w.className = 'elixir-ai-note warn';
                w.textContent = tf('goalNotMet', [res.verification.problems.join('; ')]).replace('{0}', res.verification.problems.join('; '));
                notes.appendChild(w);
            }
            me.setState('review');
            me.historyIndex = me.host.historyIndex();
            var tracked = res.ctx && res.ctx.revisionsAfter && res.ctx.revisionsAfter.count > (res.ctx.revisionsBefore ? res.ctx.revisionsBefore.count : 0);
            me.setStatus(tracked ? t('done') : t('doneUntracked'));
            me.setActions([
                { id: 'discard', label: t('discard'), fn: me.discard.bind(me) },
                { id: 'accept', label: t('accept'), primary: true, fn: me.accept.bind(me) }
            ]);
            var acc = me.el.querySelector('[data-act=accept]');
            acc && acc.focus();
        });
    };

    // Entry points: open the bar with an instruction and run it
    CommandBar.prototype.ask = function (text, autoSubmit) {
        var me = this;
        // the server-managed configuration may still be loading: wait for it before deciding
        if (!ElixirAI.Settings.isConfigured() && !this.managedLoaded) {
            this.managedLoaded = true;
            ElixirAI.Settings.loadManaged().then(function () { me.managedChecked = true; me.ask(text, autoSubmit); });
            return;
        }
        this.open();
        if (!this.input || this.input.disabled || this.state === 'running') return;
        this.input.value = text || '';
        if (autoSubmit !== false && text && ElixirAI.Settings.isConfigured()) this.submit();
        else { var me = this; setTimeout(function () { me.input.focus(); me.input.setSelectionRange(me.input.value.length, me.input.value.length); }, 0); }
    };

    CommandBar.prototype.stop = function () {
        if (this.abort) this.abort.abort();
    };

    CommandBar.prototype.finishReview = function () {
        this.setState('idle');
        this.result = null;
    };

    CommandBar.prototype.accept = function () {
        var me = this, res = this.result;
        if (!res) return;
        this.host.toolsetFactory().accept(res.ctx).then(function () {
            me.journalOutcome(res, 'accepted');
            me.setStatus(t('accepted'));
            me.finishReview();
            me.setActions([{ label: t('close'), fn: me.close.bind(me) }]);
            setTimeout(function () { me.close(); }, 600);
        });
    };

    CommandBar.prototype.discard = function () {
        var me = this, res = this.result;
        if (!res) return;
        var ts = this.host.toolsetFactory();
        // One undo restores the document as it was before the action, untracked layout included,
        // as long as nothing was edited since. Otherwise only the tracked changes can be rejected.
        var sameHistory = this.host.historyIndex() === this.historyIndex;
        if (!sameHistory && ts.canRejectAfterEdit === false) {
            me.setStatus(me.host.toolsetId==='slide' ? ElixirAI.slideText('staleDiscard') : ElixirAI.cellText('staleDiscard'), true);
            me.setActions([{ label:t('close'), fn:me.close.bind(me) },{ id:'accept',label:t('accept'),primary:true,fn:me.accept.bind(me) }]);
            return;
        }
        (sameHistory ? ts.undo() : ts.reject(res.ctx)).then(function () {
            me.journalOutcome(res, sameHistory ? 'discarded' : 'discarded-reject');
            me.setStatus(sameHistory ? t('discarded') : t('discardedReject'));
            me.finishReview();
            me.setActions([{ label: t('close'), fn: me.close.bind(me) }]);
            setTimeout(function () { me.close(); }, sameHistory ? 600 : 3000);
        });
    };

    CommandBar.prototype.journalOutcome = function (res, outcome) {
        if (!res.entry || !res.entry.id) return;
        res.entry.outcome = outcome;
        ElixirAI.Journal.update(res.entry);
    };

    ElixirAI.CommandBar = CommandBar;

    // ---------------------------------------------------------------------------------------------
    // Settings and journal dialog
    // ---------------------------------------------------------------------------------------------

    // ---------------------------------------------------------------------------------------------
    // AI connection form (File > AI page and the settings dialog): provider presets, endpoint, model,
    // key kept in this browser, where-the-data-goes notice, Test connection.
    // ---------------------------------------------------------------------------------------------

    var PRESETS = {
        custom: { url: '', model: '' },
        openai: { url: 'https://api.openai.com/v1', model: 'gpt-4.1-mini' },
        mistral: { url: 'https://api.mistral.ai/v1', model: 'mistral-small-latest' },
        ollama: { url: 'http://localhost:11434/v1', model: 'llama3.1', cors: true }
    };

    function hostOf(url) { try { return new URL(url).host; } catch (e) { return ''; } }

    function presetOf(url) {
        for (var k in PRESETS) if (k !== 'custom' && PRESETS[k].url === String(url || '').replace(/\/+$/, '')) return k;
        return 'custom';
    }

    // The sentence that says where the text goes
    ElixirAI.dataNotice = function (endpoint) {
        var S = ElixirAI.Settings;
        if (endpoint) {
            var h = hostOf(endpoint);
            return h ? ElixirAI.tf('noticeUser', { host: h }) : '';
        }
        if (S.managed) return S.managed.host ? ElixirAI.tf('noticeServer', { host: S.managed.host }) : t('noticeServerNoHost');
        return t('noticeOff');
    };

    // Preserve translations applied before this AMD module is loaded.
    var desktopStrings = ElixirAI.ConfigForm || {};
    ElixirAI.ConfigForm = {
        desktopKeyLabel: desktopStrings.desktopKeyLabel || 'API key (encrypted by your operating system)',
        desktopPrivacy: desktopStrings.desktopPrivacy || 'Your operating system encrypts the key. It is stored only on this computer, never shown again, and sent only to the AI service above, never logged.',
        // container: DOM element; host: {api, onSettingsChanged}; the form lives in container
        render: function (container, host) {
            injectStyles();
            var s = ElixirAI.Settings.get(), saved = ElixirAI.Settings.hasUserConfig(), managedOn = !!ElixirAI.Settings.managed && !saved;
            var preset = presetOf(saved ? s.endpoint : '');
            container.innerHTML =
                '<div class="elixir-ai-cfg">' +
                '<p class="lead">' + esc(t('aiLead')) + '</p>' +
                (ElixirAI.Settings.managed ? '<p class="elixir-ai-managed">' + esc(saved ? t('managedOverridden') : t('managedNote')) + '</p>' : '') +
                '<label for="elixir-ai-preset">' + esc(t('provider')) + '</label>' +
                '<select id="elixir-ai-preset">' + ['custom', 'openai', 'mistral', 'ollama'].map(function (k) {
                    return '<option value="' + k + '"' + (k === preset ? ' selected' : '') + '>' + esc(t('preset_' + k)) + '</option>';
                }).join('') + '</select>' +
                '<label for="elixir-ai-endpoint">' + esc(t('endpoint')) + '</label>' +
                '<input id="elixir-ai-endpoint" type="url" placeholder="https://api.openai.com/v1" value="' + esc(saved ? s.endpoint : '') + '">' +
                '<label for="elixir-ai-model">' + esc(t('model')) + '</label>' +
                '<input id="elixir-ai-model" list="elixir-ai-models" value="' + esc(saved || !managedOn ? s.model : '') + '">' +
                '<datalist id="elixir-ai-models"><option value="gpt-4.1-mini"><option value="gpt-5.4-mini"><option value="gpt-5.4"><option value="mistral-small-latest"><option value="llama3.1"></datalist>' +
                '<label for="elixir-ai-key">' + esc(window.eoDesktopAI ? ElixirAI.ConfigForm.desktopKeyLabel : t('keyLabel')) + '</label>' +
                '<input id="elixir-ai-key" type="password" autocomplete="off" value="' + esc(s.apiKey) + '"' + (window.eoDesktopAI && s.hasKey ? ' placeholder="••••••••"' : '') + '>' +
                '<div class="hint elixir-ai-cors" style="display:none">' + esc(t('corsNote')) + '</div>' +
                '<details class="elixir-ai-adv"><summary>' + esc(t('advanced')) + '</summary>' +
                '<label for="elixir-ai-extra">' + esc(t('extra')) + '</label>' +
                '<input id="elixir-ai-extra" placeholder=\'{"reasoning_effort":"low"}\' value="' + esc(s.extraParams) + '">' +
                '<label for="elixir-ai-rounds">' + esc(t('rounds')) + '</label>' +
                '<input id="elixir-ai-rounds" type="number" min="1" max="20" value="' + esc(s.maxRounds) + '" style="width:90px"></details>' +
                '<div class="elixir-ai-notice" id="elixir-ai-notice" role="status"></div>' +
                '<div class="hint">' + esc(window.eoDesktopAI ? ElixirAI.ConfigForm.desktopPrivacy : t('privacy')) + '</div>' +
                '<div class="elixir-ai-actions" style="justify-content:flex-start">' +
                    '<button class="elixir-ai-btn primary" data-act="save">' + esc(t('save')) + '</button>' +
                    '<button class="elixir-ai-btn" data-act="test">' + esc(t('test')) + '</button>' +
                    '<button class="elixir-ai-btn" data-act="clear">' + esc(t('clear')) + '</button>' +
                    '<span class="msg" role="status"></span>' +
                '</div></div>';
            var $ = function (sel) { return container.querySelector(sel); };
            var msg = function (text, err) { var m = $('.msg'); m.textContent = text; m.classList.toggle('error', !!err); };
            var ep = $('#elixir-ai-endpoint'), mod = $('#elixir-ai-model');
            var refresh = function () {
                $('#elixir-ai-notice').textContent = ElixirAI.dataNotice(ep.value.trim());
                $('.elixir-ai-cors').style.display = (!window.eoDesktopAI && (/^https?:\/\/(localhost|127\.0\.0\.1|\[::1\])(:|\/|$)/i.test(ep.value.trim()) || presetOf(ep.value.trim()) === 'ollama')) ? '' : 'none';
            };
            $('#elixir-ai-preset').addEventListener('change', function () {
                var p = PRESETS[this.value] || PRESETS.custom;
                var known = Object.keys(PRESETS).map(function (k) { return PRESETS[k].model; });
                ep.value = p.url;
                if (!mod.value.trim() || known.indexOf(mod.value.trim()) >= 0) mod.value = p.model;
                refresh(); ep.focus();
            });
            ep.addEventListener('input', function () { $('#elixir-ai-preset').value = presetOf(ep.value.trim()); refresh(); });
            refresh();
            var read = function () {
                var n = {
                    endpoint: ep.value.trim(), apiKey: $('#elixir-ai-key').value.trim(),
                    model: mod.value.trim(), extraParams: $('#elixir-ai-extra').value.trim(),
                    maxRounds: Math.max(1, Math.min(20, parseInt($('#elixir-ai-rounds').value, 10) || 8)),
                    prices: ElixirAI.Settings.get().prices,
                    mode: ElixirAI.Settings.get().mode, earlyStop: ElixirAI.Settings.get().earlyStop
                };
                if (n.endpoint && !ElixirAI.Settings.endpointAllowed(n.endpoint, false)) { msg(t('invalidUrl'), true); return null; }
                if (n.endpoint && !n.model) { msg(t('modelRequired'), true); return null; }
                if (n.extraParams) {
                    try { var j = JSON.parse(n.extraParams); if (!j || typeof j !== 'object' || Array.isArray(j)) throw 0; }
                    catch (e) { msg(t('invalidJson'), true); return null; }
                }
                return n;
            };
            var changed = function () { host && host.onSettingsChanged && host.onSettingsChanged(); };
            $('[data-act=save]').addEventListener('click', function () {
                var n = read();
                if (!n) return;
                Promise.resolve(!n.endpoint ? ElixirAI.Settings.clear() : ElixirAI.Settings.save(n)).then(function (saved) {
                    if (saved === false) { msg(t('notConfigured'), true); return; }
                    msg(t('saved')); changed();
                }).catch(function () { msg(t('invalidUrl'), true); });
            });
            $('[data-act=test]').addEventListener('click', function () {
                var n = read();
                if (!n) return;
                if (!n.endpoint && ElixirAI.Settings.managed) n = ElixirAI.Settings.get();
                if (!ElixirAI.Settings.isConfigured(n)) { msg(t('notConfigured'), true); return; }
                msg(t('testing'));
                ElixirAI.Connector.chat({ settings: n, messages: [{ role: 'user', content: 'Reply with: ok' }] }).then(function (r) {
                    msg(t('testOk') + ' (' + Math.round(r.ms) + ' ms)');
                }, function (e) {
                    var m = String((e && e.message) || e);
                    if (!(e && e.status) && /failed to fetch|networkerror|load failed|network request/i.test(m))
                        m = ElixirAI.tf('netError', { host: hostOf(n.endpoint) || n.endpoint });
                    msg(m, true);
                });
            });
            $('[data-act=clear]').addEventListener('click', function () {
                Promise.resolve(ElixirAI.Settings.clear()).then(function () {
                    return ElixirAI.Settings.loadManaged(true);
                }).then(function () { ElixirAI.ConfigForm.render(container, host); changed(); });
            });
        }
    };

    function SettingsDialog(host) { this.host = host; }

    SettingsDialog.prototype.open = function (focusJournal) {
        injectStyles();
        this.close();
        setGuidedTipsSuspended(this, true);
        var me = this;
        if (this.host.api) this.host.api.asc_enableKeyEvents(false);
        var mask = this.mask = document.createElement('div');
        mask.className = 'elixir-ai-modal-mask';
        var dlg = this.dlg = document.createElement('div');
        dlg.className = 'elixir-ai-dlg';
        dlg.id = 'elixir-ai-settings';
        dlg.setAttribute('role', 'dialog');
        dlg.innerHTML =
            '<h2>' + esc(t('dlgTitle')) + '</h2>' +
            '<div class="cfg-slot"></div>' +
            '<h3 id="elixir-ai-journal">' + esc(t('journalTitle')) + '</h3>' +
            '<div class="elixir-ai-actions" style="justify-content:flex-start;margin-top:0">' +
                '<button class="elixir-ai-btn" data-act="export">' + esc(t('exportJ')) + '</button>' +
                '<button class="elixir-ai-btn" data-act="clearj">' + esc(t('clearJ')) + '</button>' +
            '</div>' +
            '<div class="journal-list"></div>' +
            '<div class="elixir-ai-actions"><button class="elixir-ai-btn" data-act="close">' + esc(t('close')) + '</button></div>';
        document.body.appendChild(mask);
        document.body.appendChild(dlg);
        mask.addEventListener('click', function () { me.close(); });
        dlg.addEventListener('keydown', function (e) { if (e.key === 'Escape') me.close(); e.stopPropagation(); });
        dlg.addEventListener('keyup', function (e) { e.stopPropagation(); });
        dlg.addEventListener('keypress', function (e) { e.stopPropagation(); });
        var $ = function (sel) { return dlg.querySelector(sel); };
        ElixirAI.ConfigForm.render($('.cfg-slot'), this.host);
        $('[data-act=export]').addEventListener('click', function () {
            ElixirAI.Journal.exportJSON().then(function (json) {
                var a = document.createElement('a');
                a.href = URL.createObjectURL(new Blob([json], { type: 'application/json' }));
                a.download = 'elixir-ai-journal-' + new Date().toISOString().slice(0, 19).replace(/[:T]/g, '-') + '.json';
                document.body.appendChild(a); a.click();
                setTimeout(function () { URL.revokeObjectURL(a.href); a.remove(); }, 1000);
            });
        });
        $('[data-act=clearj]').addEventListener('click', function () { ElixirAI.Journal.clear().then(function () { me.renderJournal(); }); });
        $('[data-act=close]').addEventListener('click', function () { me.close(); });
        this.renderJournal();
        setTimeout(function () {
            if (focusJournal) $('#elixir-ai-journal').scrollIntoView();
            else $('#elixir-ai-endpoint').focus();
        }, 0);
    };

    SettingsDialog.prototype.renderJournal = function () {
        var box = this.dlg && this.dlg.querySelector('.journal-list');
        if (!box) return;
        ElixirAI.Journal.list().then(function (all) {
            if (!all.length) { box.innerHTML = '<div class="hint">' + esc(t('emptyJ')) + '</div>'; return; }
            var rows = all.slice(0, 100).map(function (e) {
                var tok = e.usage ? (e.usage.prompt_tokens + ' / ' + e.usage.completion_tokens) : '';
                var cost = (e.costUsd !== null && e.costUsd !== undefined) ? '$' + e.costUsd.toFixed(5) : '–';
                var tools = [];
                (e.rounds || []).forEach(function (r) { (r.tool_calls || []).forEach(function (c) { tools.push(c.name); }); });
                return '<tr class="entry" data-id="' + e.id + '"><td>' + esc((e.ts || '').replace('T', ' ').slice(0, 19)) + '</td><td>' + esc(e.model) +
                    '</td><td>' + esc(e.prompt) + '<div class="hint">' + esc(tools.join(', ')) + '</div></td><td>' + esc(e.outcome) + '</td><td>' + ((e.ms / 1000) || 0).toFixed(1) + ' s</td><td>' + tok + '</td><td>' + cost + '</td></tr>' +
                    '<tr class="detail" data-for="' + e.id + '" style="display:none"><td colspan="7"><pre>' + esc(JSON.stringify(e, null, 1)) + '</pre></td></tr>';
            }).join('');
            box.innerHTML = '<table><thead><tr><th>' + [t('when'), t('model'), t('prompt'), t('outcome'), t('time'), t('tokens'), t('cost')].map(esc).join('</th><th>') + '</th></tr></thead><tbody>' + rows + '</tbody></table>';
            Array.prototype.forEach.call(box.querySelectorAll('tr.entry'), function (tr) {
                tr.addEventListener('click', function () {
                    var d = box.querySelector('tr.detail[data-for="' + tr.getAttribute('data-id') + '"]');
                    d.style.display = d.style.display === 'none' ? '' : 'none';
                });
            });
        });
    };

    SettingsDialog.prototype.close = function () {
        if (this.dlg) { this.dlg.remove(); this.dlg = null; }
        if (this.mask) { this.mask.remove(); this.mask = null; }
        setGuidedTipsSuspended(this, false);
        if (this.host.api && !(this.host.bar && this.host.bar.isOpen())) this.host.api.asc_enableKeyEvents(true);
    };

    ElixirAI.SettingsDialog = SettingsDialog;

    return ElixirAI;
});

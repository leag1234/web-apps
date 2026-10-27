/*
 * SPDX-FileCopyrightText: 2026 ElixirOffice contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 *
 * Common.Controllers.ElixirAI: wires the Elixir AI layer into an editor.
 *  - adds the "Elixir" ribbon tab (Ask Elixir, AI settings, AI journal);
 *  - binds the command bar shortcut: Alt+I (Word's Copilot shortcut), Ctrl+Option+I on macOS,
 *    following the editor convention for Alt shortcuts on Mac (alt+f -> ctrl+alt+f);
 *  - hosts the command bar and the settings dialog.
 * The editor-specific tool layer is registered in Common.ElixirAI.toolsets (see
 * documenteditor/main/app/elixir-ai/WordTools.js); without a toolset nothing is added.
 */

define([
    'core',
    'common/main/lib/elixir-ai/Core',
    'common/main/lib/elixir-ai/Icons',
    'common/main/lib/elixir-ai/Strings',
    'common/main/lib/elixir-ai/View',
    'common/main/lib/elixir-ai/Setup'
], function () {
    'use strict';

    var ElixirAI = window.Common.ElixirAI;

    Common.Controllers = Common.Controllers || {};

    Common.Controllers.ElixirAI = Backbone.Controller.extend({
        toolsetId: 'word',

        initialize: function () {
            var me = this;
            this.addListeners({
                'Toolbar': {
                    'render:before': function (toolbar) { me.onToolbarRender(toolbar); },
                    // the tab strip exists only after Mixtbar.afterRender, which runs right after this event
                    'render:after': function () {
                        me.hasTab && setTimeout(function () { Common.NotificationCenter.trigger('tab:visible', 'elixir', me.isAvailable()); }, 0);
                    }
                }
            });
            Common.NotificationCenter.on('app:ready', this.onAppReady.bind(this));
        },

        onLaunch: function () {},

        appOptions: function () {
            var main = this.getApplication().getController('Main');
            return (main && main.appOptions) || {};
        },

        isAvailable: function () {
            var o = this.appOptions();
            return !!(o.isEdit && !o.isEditMailMerge && !o.isEditDiagram && !o.isEditOle && ElixirAI.toolsets && ElixirAI.toolsets[this.toolsetId]);
        },

        // AI is usable: the editor has the toolset and a model is configured (by the server or by this user).
        // Entry points remain discoverable; unconfigured actions open setup before any request.
        isActive: function () {
            return this.isAvailable() && ElixirAI.Settings.isConfigured();
        },

        // Called when the configuration may have changed (server answer known, settings saved or removed)
        refreshState: function () {
            var active = this.isActive();
            if (this.hasTab) Common.NotificationCenter.trigger('tab:visible', 'elixir', this.isAvailable());
            Common.NotificationCenter.trigger('elixir:ai-state', active, this);
        },

        // File > AI page
        openSetup: function () {
            if (!this.isAvailable()) return;
            if (!this.isActive()) {
                if (!this.setupAssistant) this.setupAssistant = new ElixirAI.SetupAssistant(this.host());
                this.setupAssistant.open(document.activeElement);
                return;
            }
            var chrome = this.getApplication().getController('ElixirChrome');
            if (chrome && chrome.openBackstage) chrome.openBackstage('elx-ai');
            else this.openSettings(false);
        },

        renderSetup: function (el) {
            ElixirAI.setLang(this.appOptions().lang);
            if (!this.api) this.api = this.getApplication().getController('Viewport').getApi();
            if (!this.isActive()) {
                el.textContent = '';
                var me = this, button = document.createElement('button');
                button.type = 'button'; button.className = 'elixir-ai-btn primary';
                button.textContent = ElixirAI.t('configure');
                button.addEventListener('click', function () { me.openSetup(); });
                el.appendChild(button);
                this.openSetup();
                return;
            }
            ElixirAI.ConfigForm.render(el, this.host());
        },

        shortcutLabel: function () { return Common.Utils.isMac ? '⌃⌥I' : 'Alt+I'; },

        onToolbarRender: function (toolbar) {
            if (!this.isAvailable()) return;
            ElixirAI.setLang(this.appOptions().lang);
            ElixirAI.injectStyles();
            var t = ElixirAI.t, me = this;
            var $panel = $('<section class="panel" data-tab="elixir" role="tabpanel" aria-labelledby="elixir">' +
                '<div class="group"><span class="btn-slot text x-huge" id="slot-btn-elixir-ask"></span></div>' +
                '<div class="separator long"></div>' +
                '<div class="group"><span class="btn-slot text x-huge" id="slot-btn-elixir-settings"></span>' +
                '<span class="btn-slot text x-huge" id="slot-btn-elixir-journal"></span></div>' +
                '</section>');
            var mk = function (slot, icon, caption, hint, fn) {
                var btn = new Common.UI.Button({
                    cls: 'btn-toolbar x-huge icon-top', iconCls: 'elixir-ai-icon elixir-ai-icon-' + icon,
                    caption: caption, hint: hint, dataHint: '1', dataHintDirection: 'bottom', dataHintOffset: 'small'
                });
                btn.render($panel.find(slot));
                btn.on('click', fn);
                return btn;
            };
            this.btnAsk = mk('#slot-btn-elixir-ask', 'sparkle', t('ask'), t('askHint') + ' (' + this.shortcutLabel() + ')', function () { me.openBar(); });
            mk('#slot-btn-elixir-settings', 'settings', t('settings'), t('settingsHint'), function () { me.openSettings(false); });
            mk('#slot-btn-elixir-journal', 'history', t('journal'), t('journal'), function () { me.openSettings(true); });
            toolbar.addTab({ action: 'elixir', caption: t('tab'), dataHintTitle: 'A', layoutname: 'toolbar-elixir' }, $panel, Common.UI.LayoutManager.lastTabIdx);
            this.hasTab = true;
        },

        onAppReady: function () {
            if (!this.isAvailable()) return;
            this.api = this.getApplication().getController('Viewport').getApi();
            var me = this, keymap = {};
            keymap[Common.Utils.isMac ? 'ctrl+alt+i' : 'alt+i'] = function (e) {
                if (Common.Utils.ModalWindow && Common.Utils.ModalWindow.isVisible()) return true;
                me.openBar();
                e && e.preventDefault && e.preventDefault();
                return false;
            };
            Common.util.Shortcuts.delegateShortcuts({ shortcuts: keymap });
            ElixirAI.setLang(this.appOptions().lang);
            ElixirAI.injectStyles();
            // the server-managed configuration (/elixir-ai/config.json), read once, before the first request
            ElixirAI.Settings.loadManaged().then(function () { me.refreshState(); });
            // the editor's other entry points (title bar search, Ask Elixir menu, "/" menu)
            var ep = ElixirAI.entryPoints && ElixirAI.entryPoints[this.toolsetId];
            ep && ep.setup(this);
            window.addEventListener('pagehide', function () {
                if (me.setupAssistant) me.setupAssistant.dispose();
                if (me.bar) me.bar.stop();
                if (me._toolset && me._toolset.dispose) me._toolset.dispose();
            }, { once: true });
            Common.NotificationCenter.trigger('elixir:ready', this);
        },

        // Run an instruction from another entry point (search box, Ask Elixir menu, "/" menu)
        ask: function (text, autoSubmit) {
            if (!this.isAvailable()) return false;
            if (!this.isActive()) { this.openSetup(); return false; }
            this.openBar();
            this.bar.ask(text, autoSubmit);
            return true;
        },

        host: function () {
            var me = this;
            if (this._host) return this._host;
            this._host = {
                toolsetId: this.toolsetId,
                api: this.api,
                toolsetFactory: function () {
                    if (!me._toolset) me._toolset = ElixirAI.toolsets[me.toolsetId].create(me.api);
                    return me._toolset;
                },
                openSettings: function () { me.openSetup(); },
                onSettingsChanged: function () {
                    me.refreshState();
                    if (me.bar && me.bar.isOpen()) me.bar.open();
                },
                // History position, to know whether Discard can simply undo the AI step
                historyIndex: function () { return (window.AscCommon && AscCommon.History) ? AscCommon.History.Index : null; },
                onClosed: function () { Common.NotificationCenter.trigger('edit:complete'); }
            };
            return this._host;
        },

        openBar: function () {
            if (!this.isAvailable()) return;
            if (!this.isActive()) { this.openSetup(); return; }
            if (!this.api) this.api = this.getApplication().getController('Viewport').getApi();
            var host = this.host();
            host.api = this.api;
            if (!this.bar) { this.bar = new ElixirAI.CommandBar(host); host.bar = this.bar; }
            this.bar.open();
        },

        openSettings: function (journal) {
            if (!journal && !this.isActive()) { this.openSetup(); return; }
            var host = this.host();
            host.api = this.api || this.getApplication().getController('Viewport').getApi();
            if (!this.dialog) this.dialog = new ElixirAI.SettingsDialog(host);
            this.dialog.open(journal);
        }
    });

    return Common.Controllers.ElixirAI;
});

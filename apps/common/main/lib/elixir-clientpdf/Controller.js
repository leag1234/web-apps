/*
 * SPDX-FileCopyrightText: 2026 ElixirOffice contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 *
 * Common.Controllers.ElixirClientPdf: the interface side of the PDF made in the browser
 * (sdkjs common/clientpdf, Word editor):
 *  - progress of long exports in the existing "Downloading…" mask;
 *  - an explicit, non-blocking message when the export had to fall back to the server, with the reason;
 *  - File > Options switch "Create PDF files in the browser" (off: the server creates every PDF).
 */

define([
    'core',
    'common/main/lib/elixir-clientpdf/Strings'
], function () {
    'use strict';

    var CP = window.Common.ElixirClientPdf;

    function esc(s) { return String(s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); }
    function engine() { return window.AscCommon && window.AscCommon.ClientPdf; }

    var CSS = [
        '.elx-cp-bar{position:fixed;z-index:1001;left:50%;bottom:44px;transform:translateX(-50%);box-sizing:border-box;width:640px;max-width:calc(100vw - 32px);',
        'display:flex;align-items:center;gap:12px;padding:10px 12px 10px 16px;background:var(--background-normal,#fff);color:var(--text-normal,#222);',
        'border:1px solid var(--border-regular-control,#ccc);border-radius:6px;box-shadow:0 4px 18px rgba(0,0,0,.22);font:12px/1.4 var(--font-family-base-custom,Arial,sans-serif);}',
        '.elx-cp-msg{flex:1 1 auto;min-width:0;}',
        '.elx-cp-btn{font:inherit;padding:5px 12px;border-radius:4px;cursor:pointer;flex:0 0 auto;border:1px solid var(--background-accent-button,#446995);',
        'background:var(--background-accent-button,#446995);color:var(--text-inverse,#fff);}',
        '#elx-cp-hint{display:block;margin-top:4px;}'
    ].join('');

    Common.Controllers.ElixirClientPdf = Backbone.Controller.extend({
        initialize: function () {
            Common.NotificationCenter.on('app:ready', this.onAppReady.bind(this));
        },

        onLaunch: function () {},

        onAppReady: function () {
            this.api = this.getApplication().getController('Viewport').getApi();
            if (!this.api || !engine()) return;      // editors without the client PDF engine
            CP.setLang(Common.Locale && Common.Locale.getCurrentLanguage && Common.Locale.getCurrentLanguage());
            var st = document.createElement('style');
            st.textContent = CSS;
            document.head.appendChild(st);
            this.api.asc_registerCallback('asc_onClientPdf', this.onEvent.bind(this));
            this.setupOptions();
        },

        onEvent: function (e) {
            if (!e) return;
            if (e.state === 'progress' && e.pages >= 4) {
                var main = this.getApplication().getController('Main');
                if (main && main.loadMask) main.loadMask.setTitle(CP.t('progress', e.page, e.pages));
            } else if (e.state === 'fallback') {
                var reason = CP.reason(e.reason, e.detail);
                this.showBar(CP.t('fallback', reason));
            }
        },

        showBar: function (text) {
            var me = this;
            this.hideBar();
            var $bar = $('<div class="elx-cp-bar" role="status"><div class="elx-cp-msg"></div><button type="button" class="elx-cp-btn"></button></div>');
            $bar.find('.elx-cp-msg').text(text);
            $bar.find('.elx-cp-btn').text(CP.t('ok')).on('click', function () { me.hideBar(); Common.NotificationCenter.trigger('edit:complete'); });
            $(document.body).append($bar);
            this.$bar = $bar;
            this.timer = setTimeout(function () { me.hideBar(); }, 15000);
        },

        hideBar: function () {
            if (this.$bar) { this.$bar.remove(); this.$bar = null; }
            clearTimeout(this.timer);
        },

        // ---- File > Options ---------------------------------------------------------------------------
        setupOptions: function () {
            var me = this, tries = 0;
            var build = function () {
                var $tb = $('#panel-settings .flex-settings table tbody').first();
                if (!$tb.length) $tb = $('#panel-settings table tbody').first();
                if (!$tb.length) return false;
                if ($('#elx-cp-chb').length) return true;
                $tb.append(
                    '<tr class="divider-group"></tr>' +
                    '<tr><td colspan="2" class="group-name"><label>' + esc(CP.t('group')) + '</label></td></tr>' +
                    '<tr><td colspan="2"><span id="elx-cp-chb"></span></td></tr>' +
                    '<tr><td colspan="2"><label id="elx-cp-hint" class="comment-text">' + esc(CP.t('hint')) + '</label></td></tr>');
                me.chb = new Common.UI.CheckBox({
                    el: $('#elx-cp-chb'),
                    labelText: CP.t('option'),
                    value: engine().isEnabled(),
                    dataHint: '1', dataHintDirection: 'left', dataHintOffset: 'small'
                });
                me.chb.on('change', function (field, newValue) { engine().setEnabled(newValue === 'checked'); });
                return true;
            };
            var timer = setInterval(function () { if (build() || ++tries > 60) clearInterval(timer); }, 500);
            $(document).on('click', '#fm-btn-settings', function () {
                setTimeout(function () { build(); me.chb && me.chb.setValue(engine().isEnabled(), true); }, 50);
            });
        }
    });

    return Common.Controllers.ElixirClientPdf;
});

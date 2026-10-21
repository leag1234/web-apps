/*
 * SPDX-FileCopyrightText: 2026 ElixirOffice contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 *
 * Common.Controllers.ElixirLocalFonts: the user-facing side of "use the fonts installed on this computer".
 *  - installs the engine hooks as soon as the editor engine exists (Core.js), before the document fonts load;
 *  - once the document is open, invites the user (non-blocking bar) when it uses fonts the server does not
 *    have, and explains the way out when the browser cannot give access (HTTP, Firefox, Safari);
 *  - adds the "Use fonts installed on this computer" switch and its status to File > Options;
 *  - draws the font box entries of local fonts with the font itself and an "On this computer" mark.
 * Shared by the Word, Excel and PowerPoint editors; no editor-specific code here.
 */

define([
    'core',
    'common/main/lib/elixir-localfonts/Core',
    'common/main/lib/elixir-localfonts/Strings'
], function () {
    'use strict';

    var LF = window.Common.ElixirLocalFonts;
    var SS_BAR = 'elixir.localfonts.bar.dismissed';

    function ssGet(k) { try { return window.sessionStorage.getItem(k); } catch (e) { return null; } }
    function ssSet(k, v) { try { window.sessionStorage.setItem(k, v); } catch (e) { /* private mode */ } }
    function esc(s) { return String(s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); }

    var CSS = [
        '.elx-lf-bar{position:fixed;z-index:1001;left:50%;bottom:44px;transform:translateX(-50%);box-sizing:border-box;width:720px;max-width:calc(100vw - 32px);',
        'display:flex;flex-wrap:wrap;align-items:center;gap:12px;padding:10px 12px 10px 16px;background:var(--background-normal,#fff);color:var(--text-normal,#222);',
        'border:1px solid var(--border-regular-control,#ccc);border-radius:6px;box-shadow:0 4px 18px rgba(0,0,0,.22);font:12px/1.4 var(--font-family-base-custom,Arial,sans-serif);}',
        '.elx-lf-msg{flex:1 1 360px;min-width:0;}',
        '.elx-lf-btns{display:flex;flex-wrap:wrap;gap:8px;flex:0 0 auto;max-width:100%;}',
        '.elx-lf-btn{font:inherit;padding:5px 12px;border-radius:4px;cursor:pointer;border:1px solid var(--border-regular-control,#ccc);background:var(--background-normal,#fff);color:var(--text-normal,#222);}',
        '.elx-lf-btn:hover{background:var(--highlight-button-hover,#e0e0e0);}',
        '.elx-lf-btn.primary{background:var(--background-accent-button,#446995);border-color:var(--background-accent-button,#446995);color:var(--text-inverse,#fff);}',
        '.elx-lf-btn.primary:hover{filter:brightness(1.1);}',
        '.elx-lf-x{flex:0 0 auto;border:0;background:none;color:var(--text-secondary,#666);font-size:18px;line-height:1;cursor:pointer;padding:2px 6px;}',
        '#elx-lf-status{display:block;margin-top:4px;}'
    ].join('');

    Common.Controllers.ElixirLocalFonts = Backbone.Controller.extend({
        initialize: function () {
            var me = this;
            LF.setLang(Common.Locale && Common.Locale.getCurrentLanguage && Common.Locale.getCurrentLanguage());
            Common.NotificationCenter.on('app:ready', this.onAppReady.bind(this));
            LF.install();   // normally done when Core.js loads
            LF.onChange(function () { me.refreshOptions(); });
            LF.thumbnail = this.thumbnail;
        },

        onLaunch: function () {},

        onAppReady: function () {
            this.api = this.getApplication().getController('Viewport').getApi();
            LF.setLang(Common.Locale && Common.Locale.getCurrentLanguage && Common.Locale.getCurrentLanguage());
            var st = document.createElement('style');
            st.textContent = CSS;
            document.head.appendChild(st);
            this.setupOptions();
            this.invite();
        },

        // ---- invitation bar ---------------------------------------------------------------------------
        namesText: function (names) {
            var shown = names.slice(0, 4).map(function (name) { return '\u2068' + name + '\u2069'; }).join(', ');
            return names.length > 4 ? shown + '…' : shown;
        },

        invite: function () {
            var me = this, needed = LF.neededNames(), sup = LF.support();
            if (!needed.length || ssGet(SS_BAR)) return;
            var base = LF.t('missing', this.namesText(needed));
            if (sup !== 'ok') {
                return this.showBar(base + ' ' + LF.t(sup), [{ text: LF.t('ok'), primary: true, click: function () { me.dismiss(); } }]);
            }
            var offer = function () {
                me.showBar(base + ' ' + LF.t('ask'), [
                    { text: LF.t('useMine'), primary: true, click: function () { me.useMine(); } },
                    { text: LF.t('notNow'), click: function () { me.dismiss(); } }
                ]);
            };
            // consent given earlier and still valid: the fonts of this computer are already in use
            if (LF.hasConsent()) LF.permission().then(function (p) { if (p !== 'granted') offer(); });
            else {
                var bridge=LF.desktopBridge();
                if(bridge && typeof bridge.fontChoice==='function')bridge.fontChoice().then(function(choice){if(choice==='unset')offer();});
                else offer();
            }
        },

        showBar: function (text, buttons) {
            this.hideBar();
            var me = this, $bar = $('<div class="elx-lf-bar" role="status"><div class="elx-lf-msg"></div><div class="elx-lf-btns"></div>' +
                '<button type="button" class="elx-lf-x" aria-label="' + esc(LF.t('close')) + '">×</button></div>');
            $bar.find('.elx-lf-msg').text(text);
            buttons.forEach(function (b) {
                $('<button type="button" class="elx-lf-btn' + (b.primary ? ' primary' : '') + '"></button>').text(b.text).on('click', b.click).appendTo($bar.find('.elx-lf-btns'));
            });
            $bar.find('.elx-lf-x').on('click', function () { me.dismiss(); });
            $(document.body).append($bar);
            this.$bar = $bar;
        },

        hideBar: function () {
            if (this.$bar) { this.$bar.remove(); this.$bar = null; }
            clearTimeout(this.barTimer);
        },

        dismiss: function () {
            var bridge=LF.desktopBridge();if(bridge && typeof bridge.fontDismiss==='function')bridge.fontDismiss().catch(function(){LF.setStatus('error');});
            ssSet(SS_BAR, '1');
            this.hideBar();
            Common.NotificationCenter.trigger('edit:complete');
        },

        // "Use my fonts": must run in the click, the browser asks its permission here
        useMine: function () {
            var me = this;
            this.hideBar();
            this.enable().then(function (res) {
                if(LF.desktopBridge())return;
                var parts = [];
                if (res.found.length) parts.push(LF.t('applied', me.namesText(res.found)));
                if (res.missing.length) parts.push(LF.t('notFound', me.namesText(res.missing)));
                var reload = res.found.length && !res.relayout;   // the editor cannot lay out again in place: reload it
                if (reload) parts.push(LF.t('reopen'));
                var buttons = [{ text: LF.t('ok'), primary: !reload, click: function () { me.dismiss(); } }];
                if (reload) buttons.unshift({ text: LF.t('reload'), primary: true, click: function () { window.location.reload(); } });
                me.showBar(parts.join(' ') || LF.t('on', LF.count), buttons);
                if (!reload) me.barTimer = setTimeout(function () { me.hideBar(); }, 7000);
            }, function (err) {
                if(LF.desktopBridge() && err && err.code==='denied')return;
                var m = LF.t(err && err.code === 'denied' ? 'denied' : (err && (err.code === 'insecure' || err.code === 'unsupported') ? err.code : 'failed'));
                me.showBar(m, [{ text: LF.t('ok'), primary: true, click: function () { me.dismiss(); } }]);
            });
        },

        enable: function () { return LF.enable(this.api); },

        // ---- File > Options ---------------------------------------------------------------------------
        setupOptions: function () {
            var me = this, tries = 0;
            var build = function () {
                var $tb = $('#panel-settings .flex-settings table tbody').first();
                if (!$tb.length) $tb = $('#panel-settings table tbody').first();
                if (!$tb.length) return false;
                if ($('#elx-lf-chb').length) return true;
                $tb.append(
                    '<tr class="divider-group"></tr>' +
                    '<tr><td colspan="2" class="group-name"><label>' + esc(LF.t('group')) + '</label></td></tr>' +
                    '<tr><td colspan="2"><span id="elx-lf-chb"></span></td></tr>' +
                    '<tr><td colspan="2"><label id="elx-lf-status" class="comment-text"></label></td></tr>' +
                    '<tr><td colspan="2"><label class="comment-text">' + esc(LF.t('hint')) + '</label></td></tr>');
                me.chb = new Common.UI.CheckBox({
                    el: $('#elx-lf-chb'),
                    labelText: LF.t('option'),
                    dataHint: '1', dataHintDirection: 'left', dataHintOffset: 'small'
                });
                me.chb.on('change', function (field, newValue) {
                    if (me.syncing) return;
                    if (newValue === 'checked') {
                        me.enable().then(function () { me.refreshOptions(); }, function () { me.refreshOptions(); });
                    } else {
                        LF.disable();
                        me.refreshOptions(LF.t('reopen'));
                    }
                });
                me.refreshOptions();
                return true;
            };
            var timer = setInterval(function () { if (build() || ++tries > 60) clearInterval(timer); }, 500);
            $(document).on('click', '#fm-btn-settings', function () { setTimeout(build, 50); });
        },

        refreshOptions: function (note) {
            if (!this.chb) return;
            var sup = LF.support(), text;
            if (sup !== 'ok') text = LF.t(sup);
            else if (LF.status === 'loading') text = LF.t('loading');
            else if (LF.status === 'on') text = LF.t('on', LF.count);
            else if (LF.status === 'denied') text = LF.t('denied');
            else if (LF.status === 'blocked' || LF.status === 'error') text = LF.t('failed');
            else text = LF.t('off');
            this.syncing = true;
            this.chb.setValue(sup === 'ok' && LF.hasConsent() && LF.status !== 'denied' && LF.status !== 'blocked' && LF.status !== 'error', true);
            this.chb.setDisabled(sup !== 'ok');
            this.syncing = false;
            $('#elx-lf-status').text(note ? text + ' ' + note : text);
        },

        // ---- font box ---------------------------------------------------------------------------------
        // Preview of a local font for the font box: its name in the font itself, and the "On this computer" mark.
        // w/h: bitmap size of the sprite cell; cssW/cssH: displayed size
        thumbnail: function (name, w, h, cssW, cssH) {
            var c = document.createElement('canvas');
            c.width = w; c.height = h;
            c.style.width = cssW + 'px'; c.style.height = cssH + 'px';
            var ctx = c.getContext('2d'), badge = LF.t('badge');
            ctx.textBaseline = 'middle';
            ctx.font = Math.round(h * 0.36) + 'px sans-serif';
            var bw = Math.min(ctx.measureText(badge).width, w * 0.4);
            ctx.save();
            ctx.beginPath(); ctx.rect(0, 0, w - bw - h * 0.5, h); ctx.clip();
            ctx.fillStyle = '#000';
            ctx.font = Math.round(h * 0.6) + 'px "' + String(name).replace(/"/g, '') + '", sans-serif';
            ctx.fillText(name, h * 0.15, h / 2);
            ctx.restore();
            ctx.fillStyle = '#808080';
            ctx.font = Math.round(h * 0.36) + 'px sans-serif';
            ctx.textAlign = 'right';
            ctx.fillText(badge, w - h * 0.2, h / 2);
            return c;
        }
    });

    return Common.Controllers.ElixirLocalFonts;
});

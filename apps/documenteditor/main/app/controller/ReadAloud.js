/*
 * SPDX-FileCopyrightText: 2026 ElixirOffice contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 *
 * DE.Controllers.ReadAloud: Review > Speech > Read Aloud.
 *
 * Reads the selected text, or the document from the paragraph of the cursor to the end, with the speech voices of
 * the operating system through the Web Speech API of the browser engine (window.speechSynthesis). Nothing leaves
 * the machine. A small bar at the top of the page holds Pause / Play, Stop, the speed and the voice (both are
 * remembered). Ctrl+Alt+Space starts and stops it, like Word.
 *
 * Notes
 *   - The text is spoken sentence by sentence: long utterances are cut off by Chromium after about 15 seconds,
 *     and pausing is done by cancelling and speaking the current sentence again (synth.pause() is not reliable
 *     on every platform).
 *   - The voice is chosen for the language at the cursor (exact locale, then same language); the bar lists the
 *     voices of that language first.
 *   - Without any voice (typical on Linux without speech-dispatcher) the ribbon button is disabled and its tooltip
 *     says why. It never throws.
 *   - The sentence being read is shown as a temporary selection in the document and kept in view. Nothing is
 *     written to the document: no formatting, no history point, so it is never marked as modified. The selection of
 *     the user is saved when reading starts and restored when it stops. The speech boundary events narrow the
 *     highlight from the chunk being spoken to its current sentence when the voice reports them. The highlight is
 *     best effort: a sentence that cannot be located in the paragraph (fields, math, unusual content) is spoken
 *     without highlight. Typing or pasting while reading stops it first, so the highlight is never overwritten.
 */

define([
    'core'
], function () {
    'use strict';

    var LOCK = 'elxNoSpeech';

    DE.Controllers.ReadAloud = Backbone.Controller.extend(_.extend({
        initialize: function () {
            this._state = 'idle';       // idle | playing | paused
            this._chunks = [];
            this._index = 0;
            this._gen = 0;              // changes whenever the speech is restarted: late events of old utterances are ignored
            this._lang = 'en-US';
            this._voices = [];
            Common.NotificationCenter.on('app:ready', this.onAppReady.bind(this));
        },

        onAppReady: function (mode) {
            var me = this;
            this.mode = mode;
            this.api = this.getApplication().getController('Viewport').getApi();
            document.addEventListener('keydown', _.bind(this.onKey, this), true);
            // a click in the document moves the selection: the original one is not given back any more
            document.addEventListener('mousedown', function (e) {
                if (me._state !== 'idle' && me._hl && !$(e.target).closest('#elx-read-aloud').length) me._saved = null, me._hlKey = null;
            }, true);
            _.each(['paste', 'cut', 'drop'], function (ev) {
                document.addEventListener(ev, function (e) {
                    if (me._state !== 'idle' && me._hl && !$(e.target).closest('#elx-read-aloud').length) me.stop();
                }, true);
            });
            var synth = this.synth();
            if (synth) {
                try {
                    synth.addEventListener ? synth.addEventListener('voiceschanged', function () { me.loadVoices(); }) : (synth.onvoiceschanged = function () { me.loadVoices(); });
                } catch (e) {}
            }
            // the list of voices arrives late in Chromium
            _.each([0, 400, 1500, 4000, 10000], function (t) { setTimeout(function () { me.loadVoices(); }, t); });
        },

        synth: function () {
            try { return window.speechSynthesis && window.SpeechSynthesisUtterance ? window.speechSynthesis : null; } catch (e) { return null; }
        },

        loadVoices: function () {
            var synth = this.synth();
            try { this._voices = synth ? Array.prototype.slice.call(synth.getVoices() || []) : []; } catch (e) { this._voices = []; }
            this.updateButton();
            this._state !== 'idle' && this.fillVoices();
        },

        available: function () {
            return this._voices.length > 0;
        },

        // ---- ribbon button ----------------------------------------------------------------------
        attachButton: function (btn) {
            this._btn = btn;
            this.updateButton();
        },

        updateButton: function () {
            var btn = this._btn;
            if (!btn) return;
            try {
                var ok = this.available();
                Common.Utils.lockControls(LOCK, !ok, {array: [btn]});
                this._hint === undefined && (this._hint = btn.options.hint);
                btn.updateHint(ok ? this._hint : this.tipNoVoice);
                btn.toggle && btn.pressed !== (this._state !== 'idle') && btn.toggle(this._state !== 'idle', true);
            } catch (e) {}
        },

        // ---- text and language ------------------------------------------------------------------
        // the selection, else the paragraphs from the one of the cursor to the end of the document
        // (or the whole document when there is no text after the cursor)
        getText: function () {
            var me = this, api = this.api, sel = '';
            this._paras = [];
            try { sel = (api.asc_GetSelectedText && api.asc_GetSelectedText(true)) || ''; } catch (e) {}
            if (/\S/.test(sel)) {
                try {
                    var ldSel = this.logicDoc(), selParas = ldSel && ldSel.GetSelectedParagraphs && ldSel.GetSelectedParagraphs();
                    selParas && selParas.length && (this._paras = _.map(selParas, function (p) { return {p: p, t: DE.Controllers.ReadAloud.paraText(p)}; }));
                } catch (e) { this._paras = []; }
                return sel;
            }
            var out = [];
            try {
                var ld = api.WordControl && api.WordControl.m_oLogicDocument;
                if (!ld) return '';
                var cur = ld.GetCurrentParagraph && ld.GetCurrentParagraph(),
                    paras = ld.GetAllParagraphs({OnlyMainDocument: true, All: true}) || [],
                    at = cur ? paras.indexOf(cur) : -1;
                if (at < 0 && cur) paras = [cur], at = 0;
                var read = function (from) {
                    var res = [];
                    me._paras = [];
                    for (var i = Math.max(0, from); i < paras.length; i++) {
                        var t = DE.Controllers.ReadAloud.paraText(paras[i]);
                        me._paras.push({p: paras[i], t: t});
                        res.push(t);
                    }
                    return res.join('\n');
                };
                out.push(read(at));
                // the cursor is after the last text: read the document from its start instead of nothing
                !/\S/.test(out[0]) && at > 0 && (out = [read(0)]);
            } catch (e) {}
            return out.join('\n');
        },

        logicDoc: function () {
            try { return this.api && this.api.WordControl && this.api.WordControl.m_oLogicDocument || null; } catch (e) { return null; }
        },

        // BCP 47 tag of the text at the cursor
        getLang: function () {
            var lcid;
            try {
                var rc = this.getApplication().getController('Common.Controllers.ReviewChanges');
                lcid = rc && parseInt(rc._textLang);
                if (!lcid && this.api.asc_GetRevealFormatting) lcid = ((this.api.asc_GetRevealFormatting() || {font: {}}).font || {}).lang;
                var info = lcid && Common.util.LanguageInfo.getLocalLanguageName(lcid);
                if (info && info[0]) return info[0];
            } catch (e) {}
            return 'en-US';
        },

        pickVoice: function (lang, name) {
            return DE.Controllers.ReadAloud.pickVoice(this._voices, lang, name);
        },

        // ---- playback ---------------------------------------------------------------------------
        toggle: function () {
            if (this._state === 'idle') this.start();
            else this.stop();
        },

        start: function () {
            var synth = this.synth();
            if (!synth || !this.available()) return this.updateButton();
            var text = this.getText(), chunks = DE.Controllers.ReadAloud.splitText(text);
            if (!chunks.length) {
                this.showBar();
                this._chunks = [];
                this.setStatus(this.txtNothing);
                this._state = 'paused';
                this.updateBar();
                return;
            }
            this.cancel();
            this._chunks = chunks;
            this._index = 0;
            this._lang = this.getLang();
            this.prepareHighlight(chunks);
            this._state = 'playing';
            this.showBar();
            this.speak();
            this.updateButton();
        },

        rate: function () {
            var r = parseFloat(Common.localStorage.getItem('de-readaloud-rate'));
            return r >= 0.5 && r <= 3 ? r : 1;
        },

        speak: function () {
            var me = this, synth = this.synth(), gen = ++this._gen, text = this._chunks[this._index];
            if (!synth || this._state !== 'playing') return;
            if (text === undefined) return this.stop();
            try {
                var u = new window.SpeechSynthesisUtterance(text),
                    voice = this.pickVoice(this._lang, Common.localStorage.getItem('de-readaloud-voice'));
                u.lang = voice ? voice.lang : this._lang;
                voice && (u.voice = voice);
                u.rate = this.rate();
                u.onstart = function () {
                    if (gen !== me._gen) return;
                    me.highlight(me._index);
                };
                u.onboundary = function (e) {
                    if (gen !== me._gen || !e || typeof e.charIndex !== 'number' || (e.name && e.name !== 'word' && e.name !== 'sentence')) return;
                    me.highlight(me._index, e.charIndex);
                };
                u.onend = function () {
                    if (gen !== me._gen || me._state !== 'playing') return;
                    me._index++;
                    me.speak();
                };
                u.onerror = function (e) {
                    if (gen !== me._gen) return;
                    if (e && (e.error === 'canceled' || e.error === 'interrupted')) return;
                    me.fail();
                };
                this._utterance = u;        // kept: Chromium drops utterances that are garbage collected while speaking
                this.updateBar();
                this.highlight(this._index);        // already now: the voice may start late, and some engines send no start event
                synth.speak(u);
            } catch (e) {
                this.fail();
            }
        },

        pause: function () {
            if (this._state !== 'playing') return;
            this._state = 'paused';
            this.cancel();
            this.updateBar();
        },

        resume: function () {
            if (this._state !== 'paused' || !this._chunks.length) return;
            this._state = 'playing';
            this.speak();
        },

        // speaks the current sentence again (new speed or voice)
        restart: function () {
            if (this._state !== 'playing') return;
            this.cancel();
            this.speak();
        },

        cancel: function () {
            this._gen++;
            try { this.synth() && this.synth().cancel(); } catch (e) {}
        },

        stop: function () {
            this._state = 'idle';
            this.cancel();
            this._chunks = [];
            this._utterance = null;
            this.clearHighlight();
            this.hideBar();
            this.updateButton();
            if (this.api) {
                this.api.asc_enableKeyEvents && this.api.asc_enableKeyEvents(true);
                Common.NotificationCenter.trigger('edit:complete');
            }
        },

        fail: function () {
            this._state = 'paused';
            this.cancel();
            this.setStatus(this.txtFailed);
            this.updateBar();
        },

        onKey: function (e) {
            var K = Common.UI.Keys;
            if (e.keyCode === K.SPACE && e.ctrlKey && e.altKey && !e.shiftKey) {
                if (Common.Utils.ModalWindow.isVisible()) return;
                e.preventDefault();
                e.stopPropagation();
                this.toggle();
            } else if (e.keyCode === K.ESC && this._state !== 'idle' && $(document.activeElement).closest('#elx-read-aloud').length) {
                e.preventDefault();
                e.stopPropagation();
                this.stop();
            } else if (this._state !== 'idle' && this._hl && DE.Controllers.ReadAloud.isEditingKey(e) && !$(e.target).closest('#elx-read-aloud').length) {
                // the highlight is a real selection: give the original one back before the key would replace the sentence
                this.stop();
            }
        },

        // ---- highlight of the sentence being read -----------------------------------------------
        // Units are the sentences of the chunks, located in the paragraphs the text was taken from.
        prepareHighlight: function (chunks) {
            var units = [], texts = [];
            this._hl = null;
            this._saved = null;
            this._units = [];
            this._byChunk = [];
            this._hlKey = this._boundary = null;
            if (!this._paras || !this._paras.length) return;
            _.each(chunks, function (c, ci) {
                var list = [];
                _.each(DE.Controllers.ReadAloud.sentences(c), function (s) {
                    list.push(units.length);
                    units.push({chunk: ci, start: s.start, text: s.text});
                    texts.push(s.text);
                });
                this._byChunk.push(list);
            }, this);
            var at = DE.Controllers.ReadAloud.locate(texts, _.map(this._paras, function (p) { return p.t; }));
            _.each(units, function (u, i) { u.at = at[i]; });
            this._units = units;
        },

        // Shows the chunk being spoken, or - with the character a speech boundary event reports - the sentence there.
        // A voice without boundary events keeps the whole chunk shown (a chunk is a few sentences at most).
        highlight: function (ci, charIndex) {
            try {
                var list = this._byChunk && this._byChunk[ci];
                if (!list || !list.length) return;
                var isBoundary = typeof charIndex === 'number', from, to, i;
                isBoundary && (this._boundary = ci);
                if (isBoundary || this._boundary === ci) {
                    from = to = this._units[list[0]];
                    for (i = 0; i < list.length; i++) if (this._units[list[i]].start <= (charIndex || 0)) from = to = this._units[list[i]];
                } else {
                    from = this._units[list[0]];
                    to = this._units[list[list.length - 1]];
                }
                if (!from.at || !to.at || from.at.para !== to.at.para) return;
                var key = from.at.para + ':' + from.at.start + ':' + to.at.end;
                if (key === this._hlKey) return;
                if (this.select(this._paras[from.at.para].p, from.at.start, to.at.end)) this._hlKey = key;
            } catch (e) {}
        },

        // makes [start, end) of the text of the paragraph the (temporary) selection and brings it into view
        select: function (para, start, end) {
            var ld = this.logicDoc();
            if (!ld || !para || (para.IsUseInDocument && !para.IsUseInDocument())) return false;
            var pos = DE.Controllers.ReadAloud.rangePositions(para, start, end);
            if (!pos) return false;
            if (!this._hl) {
                this._saved = ld.GetSelectionState();
                this._hl = true;
            }
            var top = pos.start[0].Class;
            top.SetThisElementCurrent && top.SetThisElementCurrent(true);
            top.SetContentPosition(pos.start, 0, 0);
            top.SetSelectionByContentPositions(pos.start, pos.end);
            ld.RecalculateCurPos();
            ld.Document_UpdateInterfaceState();
            ld.Document_UpdateSelectionState();
            ld.Document_UpdateRulersState();
            return true;
        },

        // gives the selection of the user back
        clearHighlight: function () {
            var ld = this.logicDoc(), saved = this._saved;
            this._hl = null;
            this._saved = null;
            this._hlKey = null;
            this._boundary = null;
            if (!ld || !saved) return;
            try {
                ld.SetSelectionState(saved);
                ld.RecalculateCurPos();
                ld.Document_UpdateInterfaceState();
                ld.Document_UpdateSelectionState();
                ld.Document_UpdateRulersState();
            } catch (e) {}
        },

        // ---- the bar ----------------------------------------------------------------------------
        ICON_STOP: '<svg class="elx-icon" width="20" height="20" viewBox="0 0 20 20" fill="currentColor" aria-hidden="true" focusable="false">' +
            '<path d="M15.5 4c.28 0 .5.22.5.5v11a.5.5 0 0 1-.5.5h-11a.5.5 0 0 1-.5-.5v-11c0-.28.22-.5.5-.5zm-11-1C3.67 3 3 3.67 3 4.5v11c0 .83.67 1.5 1.5 1.5h11c.83 0 1.5-.67 1.5-1.5v-11c0-.83-.67-1.5-1.5-1.5z"/></svg>',

        icon: function (name) {
            return window.ElixirChromeIcons && window.ElixirChromeIcons.has(name) ? window.ElixirChromeIcons.svg(name, 20) : '';
        },

        showBar: function () {
            var me = this;
            if (!$('#elx-read-aloud-style').length) {
                $('<style id="elx-read-aloud-style"></style>').text(
                    '#elx-read-aloud{position:absolute;top:8px;left:50%;transform:translateX(-50%);z-index:510;display:flex;align-items:center;gap:6px;' +
                    'padding:4px 6px;background:var(--background-toolbar,#fff);color:var(--text-normal,#333);border:1px solid var(--border-toolbar,#ccc);' +
                    'border-radius:6px;box-shadow:0 2px 8px rgba(0,0,0,.18);font-size:11px;max-width:calc(100% - 24px);}' +
                    '#elx-read-aloud button{display:flex;align-items:center;justify-content:center;width:28px;height:28px;border:0;border-radius:4px;background:none;color:inherit;cursor:pointer;padding:0;}' +
                    '#elx-read-aloud button:hover{background:var(--highlight-button-hover,#e0e0e0);}' +
                    '#elx-read-aloud button:focus-visible,#elx-read-aloud select:focus-visible{outline:2px solid var(--border-control-focus,#2a6ebb);}' +
                    '#elx-read-aloud select{height:24px;max-width:170px;font-size:11px;color:inherit;background:var(--background-normal,#fff);border:1px solid var(--border-regular-control,#ccc);border-radius:3px;}' +
                    '#elx-read-aloud .elx-ra-status{min-width:44px;text-align:center;color:var(--text-secondary,#666);white-space:nowrap;}' +
                    '#elx-read-aloud .elx-ra-sep{width:1px;height:20px;background:var(--border-divider,#e0e0e0);}'
                ).appendTo('head');
            }
            this.hideBar();
            var $host = $('#editor_sdk').parent(),
                $bar = $('<div id="elx-read-aloud" role="toolbar" tabindex="-1">' +
                    '<button type="button" class="elx-ra-play"></button><button type="button" class="elx-ra-stop"></button>' +
                    '<span class="elx-ra-status" role="status" aria-live="polite"></span><span class="elx-ra-sep"></span>' +
                    '<select class="elx-ra-rate"></select><select class="elx-ra-voice"></select>' +
                    '<button type="button" class="elx-ra-close"></button></div>');
            $bar.attr('aria-label', this.txtReadAloud);
            $bar.find('.elx-ra-stop').html(this.ICON_STOP).attr({title: this.txtStop, 'aria-label': this.txtStop}).on('click', function () { me.stop(); });
            $bar.find('.elx-ra-close').html(this.icon('dismiss')).attr({title: this.txtClose, 'aria-label': this.txtClose}).on('click', function () { me.stop(); });
            $bar.find('.elx-ra-play').on('click', function () {
                if (me._state === 'playing') me.pause();
                else if (me._state === 'paused' && me._chunks.length) { me.resume(); me.updateBar(); }
                else me.start();
            });
            var $rate = $bar.find('.elx-ra-rate').attr({title: this.txtSpeed, 'aria-label': this.txtSpeed}), cur = this.rate();
            _.each([0.5, 0.75, 1, 1.25, 1.5, 2], function (r) { $('<option>').val(r).text(r + 'x').prop('selected', r === cur).appendTo($rate); });
            $rate.on('change', function () {
                Common.localStorage.setItem('de-readaloud-rate', $rate.val());
                me.restart();
            });
            $bar.find('.elx-ra-voice').attr({title: this.txtVoice, 'aria-label': this.txtVoice}).on('change', function () {
                Common.localStorage.setItem('de-readaloud-voice', $(this).val());
                me.restart();
            });
            $host.css('position') === 'static' && $host.css('position', 'relative');
            $host.append($bar);
            this._status = '';
            this.fillVoices();
            this.updateBar();
        },

        hideBar: function () {
            $('#elx-read-aloud').remove();
        },

        // the voices of the language of the text first
        fillVoices: function () {
            var $sel = $('#elx-read-aloud .elx-ra-voice');
            if (!$sel.length) return;
            var prim = String(this._lang).split('-')[0].toLowerCase(),
                chosen = this.pickVoice(this._lang, Common.localStorage.getItem('de-readaloud-voice')),
                same = [], other = [];
            _.each(this._voices, function (v) { (String(v.lang).split(/[-_]/)[0].toLowerCase() === prim ? same : other).push(v); });
            $sel.empty();
            _.each(same.concat(other), function (v) {
                $('<option>').val(v.name).text(v.name + ' (' + v.lang + ')').prop('selected', v === chosen).appendTo($sel);
            });
        },

        setStatus: function (text) {
            this._status = text;
            $('#elx-read-aloud .elx-ra-status').text(text);
        },

        updateBar: function () {
            var $bar = $('#elx-read-aloud');
            if (!$bar.length) return;
            var playing = this._state === 'playing', n = this._chunks.length;
            $bar.find('.elx-ra-play').html(this.icon(playing ? 'pause' : 'play')).attr({title: playing ? this.txtPause : this.txtPlay, 'aria-label': playing ? this.txtPause : this.txtPlay})
                .prop('disabled', !n && !playing);
            if (n) this.setStatus(Math.min(this._index + 1, n) + ' / ' + n);
        },

        txtReadAloud: 'Read Aloud',
        tipNoVoice: 'Read Aloud needs a speech voice, and this computer has none. On Linux, install speech-dispatcher and a voice such as espeak-ng, then restart ElixirOffice.',
        txtPlay: 'Play',
        txtPause: 'Pause',
        txtStop: 'Stop',
        txtClose: 'Close',
        txtSpeed: 'Reading speed',
        txtVoice: 'Voice',
        txtNothing: 'No text to read',
        txtFailed: 'Reading stopped'
    }, DE.Controllers.ReadAloud || {}));

    // ---- pure helpers (also used by the unit tests) ------------------------------------------------
    // Cuts a text into sentences of at most `max` characters, in order, without empty ones.
    DE.Controllers.ReadAloud.splitText = function (text, max) {
        max = max || 200;
        var out = [], parts = String(text || '').replace(/\r\n?/g, '\n').split(/\n+/);
        for (var p = 0; p < parts.length; p++) {
            var para = parts[p].replace(/\s+/g, ' ').trim();
            if (!para) continue;
            var sentences = para.match(/[^.!?…。！？؟]+(?:[.!?…。！？؟]+["'”»)\]]*|$)\s*/g) || [para], cur = '';
            for (var s = 0; s < sentences.length; s++) {
                var sent = sentences[s].trim();
                if (!sent) continue;
                if (sent.length > max) {
                    cur && out.push(cur), cur = '';
                    var words = sent.split(' '), line = '';
                    for (var w = 0; w < words.length; w++) {
                        if (line && (line + ' ' + words[w]).length > max) out.push(line), line = words[w];
                        else line = line ? line + ' ' + words[w] : words[w];
                    }
                    line && out.push(line);
                } else if (cur && (cur + ' ' + sent).length > max) {
                    out.push(cur);
                    cur = sent;
                } else {
                    cur = cur ? cur + ' ' + sent : sent;
                }
            }
            cur && out.push(cur);
        }
        return out;
    };

    // Sentences of a chunk (as speech reports positions inside the chunk): [{start, text}]
    DE.Controllers.ReadAloud.sentences = function (chunk) {
        var re = /[^.!?…。！？؟]+(?:[.!?…。！？؟]+["'”»)\]]*|$)\s*/g, out = [], m;
        chunk = String(chunk || '');
        while ((m = re.exec(chunk))) {
            if (!m[0]) { re.lastIndex++; continue; }
            /\S/.test(m[0]) && out.push({start: m.index + (m[0].length - m[0].replace(/^\s+/, '').length), text: m[0].trim()});
        }
        return out.length ? out : (/\S/.test(chunk) ? [{start: 0, text: chunk.trim()}] : []);
    };

    // The text of a paragraph, the way it is read (and measured for the highlight)
    DE.Controllers.ReadAloud.paraText = function (para) {
        return para.GetText({Numbering: false, ParaSeparator: '', NewLineSeparator: '\n', TabSymbol: ' '});
    };

    // Finds the texts, in order, in the texts of the paragraphs, ignoring white space: [{para, start, end}|null]
    // (start and end are offsets in the text of the paragraph, end excluded)
    DE.Controllers.ReadAloud.locate = function (units, paraTexts) {
        var maps = [], pi = 0, pos = 0, out = [];
        var squeeze = function (i) {
            if (!maps[i]) {
                var t = String(paraTexts[i] || ''), s = '', m = [];
                for (var k = 0; k < t.length; k++) if (!/\s/.test(t.charAt(k))) { s += t.charAt(k); m.push(k); }
                maps[i] = {s: s, m: m};
            }
            return maps[i];
        };
        for (var u = 0; u < units.length; u++) {
            var key = String(units[u] || '').replace(/\s+/g, ''), found = null;
            if (key) {
                for (var p = pi, from = pos; p < paraTexts.length; p++, from = 0) {
                    var e = squeeze(p), at = e.s.indexOf(key, from);
                    if (at >= 0) {
                        found = {para: p, start: e.m[at], end: e.m[at + key.length - 1] + 1};
                        pi = p;
                        pos = at + key.length;
                        break;
                    }
                }
            }
            out.push(found);
        }
        return out;
    };

    // Document positions (the form Select of the builder API uses) of [start, end) in the text of the paragraph,
    // counted like paraText does: {start: [...], end: [...]} or null
    DE.Controllers.ReadAloud.rangePositions = function (para, start, end) {
        var offset = 0, sPos = null, ePos = null;
        var acc = {Text: '', BreakOnNonText: false, ParaSeparator: '', NewLineSeparator: '\n', TabSymbol: ' ', TableCellSeparator: '', TableRowSeparator: ''},
            noPara = function () { return null; };
        if (!para || !para.CheckRunContent || end <= start) return null;
        para.CheckRunContent(function (run) {
            var items = run && run.Content;
            if (!items || !run.Get_Text) return false;
            for (var i = 0; i < items.length; i++) {
                acc.Text = '';
                run.Get_Text.call({Content: [items[i]], GetParagraph: noPara}, acc);
                var len = acc.Text.length;
                if (!len) continue;
                if (!sPos && start < offset + len) {
                    sPos = run.GetDocumentPositionFromObject();
                    sPos.push({Class: run, Position: i});
                }
                if (sPos && end <= offset + len) {
                    ePos = run.GetDocumentPositionFromObject();
                    ePos.push({Class: run, Position: i + 1});
                    return true;
                }
                offset += len;
            }
            return false;
        });
        return sPos && ePos ? {start: sPos, end: ePos} : null;
    };

    // keys that would change the document (or the selection that holds the highlight)
    DE.Controllers.ReadAloud.isEditingKey = function (e) {
        var k = e.keyCode, mod = e.ctrlKey || e.metaKey;
        if (k === 8 || k === 46 || k === 13 || k === 9) return true;
        if (typeof e.key === 'string' && e.key.length === 1 && (!mod || e.altKey)) return true;
        return !!mod && !e.altKey && /^(x|v|z|y|b|i|u|d|k|a)$/i.test(e.key || '');
    };

    // the voice with that name when it exists and speaks the language, else one for the exact locale, else for the same language, else none
    DE.Controllers.ReadAloud.pickVoice = function (voices, lang, name) {
        voices = voices || [];
        var norm = function (l) { return String(l || '').replace(/_/g, '-').toLowerCase(); },
            want = norm(lang), prim = want.split('-')[0], i;
        if (name) for (i = 0; i < voices.length; i++) if (voices[i].name === name && norm(voices[i].lang).split('-')[0] === prim) return voices[i];
        for (i = 0; i < voices.length; i++) if (norm(voices[i].lang) === want) return voices[i];
        for (i = 0; i < voices.length; i++) if (norm(voices[i].lang).split('-')[0] === prim) return voices[i];
        return null;
    };

    return DE.Controllers.ReadAloud;
});

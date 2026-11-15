// SPDX-FileCopyrightText: 2026 ElixirOffice contributors
// SPDX-License-Identifier: AGPL-3.0-only
'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const root = path.resolve(__dirname, '../..');
const main = path.join(root, 'apps/documenteditor/main');

// the sentence a utterance speaks, newest last
function load(options = {}) {
    const spoken = [], calls = [];
    const synth = {
        speak(u) { spoken.push(u); },
        cancel() { calls.push('cancel'); },
        getVoices() { return options.voices || []; },
        addEventListener() {}
    };
    const store = {};
    const context = {
        DE: {Controllers: {}},
        Backbone: {Controller: {extend: props => { const C = function () {}; Object.assign(C.prototype, props); return C; }}},
        _: {
            extend: Object.assign, each: (a, f, o) => a.forEach(f, o), map: (a, f) => a.map(f), bind: (f, o) => f.bind(o), isFunction: f => typeof f === 'function'
        },
        Common: {
            NotificationCenter: {on() {}, trigger() {}},
            localStorage: {getItem: k => store[k] || null, setItem: (k, v) => { store[k] = v; }},
            Utils: {lockControls() {}, ModalWindow: {isVisible: () => false}},
            UI: {Keys: {SPACE: 32, ESC: 27}}
        },
        window: {speechSynthesis: synth, SpeechSynthesisUtterance: function (t) { this.text = t; }},
        document: {addEventListener() {}},
        $: () => new Proxy({length: 0}, {get: (t, k) => (k in t ? t[k] : function () { return this; })}),
        setTimeout: () => 0,
        console
    };
    context.window.window = context.window;
    context.define = (_deps, factory) => factory();
    vm.runInNewContext(fs.readFileSync(path.join(main, 'app/controller/ReadAloud.js'), 'utf8'), context, {filename: 'ReadAloud.js'});
    const C = context.DE.Controllers.ReadAloud;
    const ctl = new C();
    ctl.initialize();
    ctl._voices = options.voices || [];
    ctl.synth = () => synth;
    ctl.getText = () => options.text || 'One. Two. Three.';
    ctl.getLang = () => options.lang || 'en-US';
    ctl.showBar = ctl.hideBar = ctl.updateBar = ctl.fillVoices = function () {};
    ctl.api = options.ld ? {WordControl: {m_oLogicDocument: options.ld}} : {};
    ctl.getApplication = () => ({getController: () => ({})});
    return {C, ctl, spoken, calls, synth};
}

const plain = v => JSON.parse(JSON.stringify(v));
const VOICES = [{name: 'Anna', lang: 'de-DE'}, {name: 'Alice', lang: 'fr_FR'}, {name: 'Sam', lang: 'en-US'}, {name: 'Kate', lang: 'en-GB'}];

test('splitText cuts sentences, keeps order and drops empty lines', () => {
    const {C} = load();
    assert.deepEqual(plain(C.splitText('Hello there. How are you?\n\n  \nFine!')), ['Hello there. How are you?', 'Fine!']);
    assert.deepEqual(plain(C.splitText('')), []);
    assert.deepEqual(plain(C.splitText(null)), []);
    assert.deepEqual(plain(C.splitText('A.\r\nB.')), ['A.', 'B.']);
});

test('splitText never produces a chunk longer than the limit', () => {
    const {C} = load();
    const long = 'word '.repeat(300).trim() + '. Short one.';
    const chunks = C.splitText(long, 100);
    assert.ok(chunks.every(c => c.length <= 100), 'chunk too long');
    assert.equal(chunks.join(' ').replace(/\s+/g, ' '), long);
});

test('pickVoice prefers the exact locale, then the language, then nothing', () => {
    const {C} = load();
    assert.equal(C.pickVoice(VOICES, 'fr-FR').name, 'Alice');
    assert.equal(C.pickVoice(VOICES, 'en-GB').name, 'Kate');
    assert.equal(C.pickVoice(VOICES, 'en-AU').name, 'Sam');
    assert.equal(C.pickVoice(VOICES, 'ja-JP'), null);
    assert.equal(C.pickVoice([], 'en-US'), null);
});

test('a remembered voice is only used for its own language', () => {
    const {C} = load();
    assert.equal(C.pickVoice(VOICES, 'en-US', 'Kate').name, 'Kate');
    assert.equal(C.pickVoice(VOICES, 'de-DE', 'Kate').name, 'Anna');
});

test('Read Aloud is unavailable without a voice and never speaks', () => {
    const {ctl, spoken} = load({voices: []});
    assert.equal(ctl.available(), false);
    ctl.start();
    assert.equal(spoken.length, 0);
    assert.equal(ctl._state, 'idle');
});

test('plays sentence by sentence, pauses, resumes the same sentence and stops', () => {
    const {ctl, spoken, calls} = load({voices: VOICES, text: 'One. Two. Three.', lang: 'fr-FR'});
    ctl.start();
    assert.equal(ctl._state, 'playing');
    assert.equal(spoken.length, 1);
    assert.equal(spoken[0].text, 'One. Two. Three.');
    assert.equal(spoken[0].voice.name, 'Alice');
    assert.equal(spoken[0].lang, 'fr_FR');
    spoken[0].onend();
    assert.equal(ctl._state, 'idle');       // the whole text fitted in one chunk: done after it

    const small = load({voices: VOICES, text: 'First paragraph.\nSecond paragraph.'});
    small.ctl.start();
    assert.equal(small.spoken.length, 1);
    small.ctl.pause();
    assert.equal(small.ctl._state, 'paused');
    assert.ok(small.calls.includes('cancel'));
    small.spoken[0].onend();                // a late event of the cancelled utterance is ignored
    assert.equal(small.spoken.length, 1);
    small.ctl.resume();
    assert.equal(small.spoken.length, 2);
    assert.equal(small.spoken[1].text, 'First paragraph.');
    small.spoken[1].onend();
    assert.equal(small.spoken.length, 3);
    assert.equal(small.spoken[2].text, 'Second paragraph.');
    small.ctl.stop();
    assert.equal(small.ctl._state, 'idle');
    small.spoken[2].onend();
    assert.equal(small.spoken.length, 3);
});

test('a speech error pauses instead of throwing', () => {
    const {ctl, spoken} = load({voices: VOICES, text: 'Hi there.'});
    ctl.start();
    spoken[0].onerror({error: 'synthesis-failed'});
    assert.equal(ctl._state, 'paused');
});

test('ribbon, bar and key tips strings exist in the 9 main locales', () => {
    const keys = ['DE.Views.RibbonTabs.txtThesaurus', 'DE.Views.RibbonTabs.tipThesaurus', 'DE.Views.RibbonTabs.txtReadAloud',
        'DE.Views.RibbonTabs.tipReadAloud', 'DE.Views.RibbonTabs.txtGroupSpeech', 'DE.Controllers.ReadAloud.tipNoVoice',
        'DE.Controllers.ReadAloud.txtPlay', 'DE.Controllers.ReadAloud.txtPause', 'DE.Controllers.ReadAloud.txtStop',
        'DE.Controllers.ReadAloud.txtSpeed', 'DE.Controllers.ReadAloud.txtVoice', 'DE.Controllers.ReadAloud.txtNothing'];
    const en = JSON.parse(fs.readFileSync(path.join(main, 'locale/en.json'), 'utf8'));
    for (const l of ['en', 'fr', 'de', 'es', 'it', 'ru', 'nl', 'pt', 'pt-pt']) {
        const d = JSON.parse(fs.readFileSync(path.join(main, 'locale', l + '.json'), 'utf8'));
        for (const k of keys) {
            assert.ok(d[k], l + ' misses ' + k);
        }
    }
    // the translations are real: the French ones differ from English
    const fr = JSON.parse(fs.readFileSync(path.join(main, 'locale/fr.json'), 'utf8'));
    assert.notEqual(fr['DE.Views.RibbonTabs.txtReadAloud'], en['DE.Views.RibbonTabs.txtReadAloud']);
    assert.equal(fr['DE.Views.RibbonTabs.txtReadAloud'], 'Lecture à voix haute');
});

test('the controller is registered in both app entry points', () => {
    for (const f of ['app.js', 'app_dev.js']) {
        const s = fs.readFileSync(path.join(main, f), 'utf8');
        assert.match(s, /,'ReadAloud'/);
        assert.match(s, /documenteditor\/main\/app\/controller\/ReadAloud/);
    }
});

// ---- highlight of the sentence being read ----------------------------------------------------------
// A tiny model of the paragraph/run structure of sdkjs: runs hold items, Get_Text turns items into text.
function makeDoc(texts) {
    const calls = [], doc = {calls, selection: null, state: {id: 'user-selection'}};
    const rec = name => (...a) => { calls.push(name); return a; };
    Object.assign(doc, {
        GetSelectionState: () => { calls.push('GetSelectionState'); return doc.state; },
        SetSelectionState: s => { calls.push('SetSelectionState'); doc.selection = s; },
        SetThisElementCurrent: rec('SetThisElementCurrent'),
        SetContentPosition: rec('SetContentPosition'),
        SetSelectionByContentPositions: (s, e) => { calls.push('select'); doc.selection = {s, e}; },
        RecalculateCurPos: rec('RecalculateCurPos'),
        Document_UpdateInterfaceState: rec('ui'), Document_UpdateSelectionState: rec('sel'), Document_UpdateRulersState: rec('ruler'),
        GetSelectedParagraphs: () => doc.selectedParas || null,
        History: new Proxy({}, {get: (t, k) => { throw new Error('history touched: ' + String(k)); }})
    });
    doc.paras = texts.map((chunks, pi) => {
        const runs = chunks.map(c => {
            const items = Array.from(c).map(ch => ({t: ch}));
            const run = {Content: items, Get_Text(acc) { for (const it of this.Content) acc.Text += it.t; },
                GetDocumentPositionFromObject: () => [{Class: doc, Position: pi}]};
            return run;
        });
        const full = chunks.join('');
        return {runs, IsUseInDocument: () => true, GetText: () => full,
            CheckRunContent(cb) { for (const r of runs) if (cb(r)) return true; return false; }};
    });
    return doc;
}

function loadHl(texts, options = {}) {
    const doc = makeDoc(texts), r = load(Object.assign({voices: VOICES, ld: doc}, options));
    r.ctl._paras = doc.paras.map(p => ({p, t: p.GetText()}));
    r.ctl.getText = () => doc.paras.map(p => p.GetText()).join('\n');
    r.doc = doc;
    return r;
}

// the text covered by the last selection, read back through the runs
function selectedText(doc) {
    const {s, e} = doc.selection, a = s[s.length - 1], b = e[e.length - 1];
    const para = doc.paras.find(p => p.runs.includes(a.Class));
    let out = '', on = false;
    for (const run of para.runs) for (let i = 0; i <= run.Content.length; i++) {
        if (run === a.Class && i === a.Position) on = true;
        if (run === b.Class && i === b.Position) return out;
        if (on && i < run.Content.length) out += run.Content[i].t;
    }
    return out;
}

test('sentences keep their offsets inside the chunk', () => {
    const {C} = load();
    assert.deepEqual(plain(C.sentences('Hello world.  Second one! Last')), [
        {start: 0, text: 'Hello world.'}, {start: 14, text: 'Second one!'}, {start: 26, text: 'Last'}]);
    assert.deepEqual(plain(C.sentences('   ')), []);
});

test('locate finds the units in order, ignoring white space, and null for what is missing', () => {
    const {C} = load();
    const at = plain(C.locate(['Hello world.', 'Second one.', 'Nothing here.', 'Third.'], ['Hello  world.\tSecond one.', 'Third.']));
    assert.deepEqual(at[0], {para: 0, start: 0, end: 12 + 1});
    assert.deepEqual(at[1], {para: 0, start: 14, end: 25});
    assert.equal(at[2], null);
    assert.deepEqual(at[3], {para: 1, start: 0, end: 6});
});

test('rangePositions crosses runs and ends exactly after the last character', () => {
    const doc = makeDoc([['Hello wor', 'ld. Second one.']]), {C} = load();
    const pos = C.rangePositions(doc.paras[0], 6, 12);
    assert.equal(pos.start[pos.start.length - 1].Class, doc.paras[0].runs[0]);
    assert.equal(pos.start[pos.start.length - 1].Position, 6);
    assert.equal(pos.end[pos.end.length - 1].Class, doc.paras[0].runs[1]);
    assert.equal(pos.end[pos.end.length - 1].Position, 3);
    doc.selection = {s: pos.start, e: pos.end};
    assert.equal(selectedText(doc), 'world.');
    assert.equal(C.rangePositions(doc.paras[0], 5, 5), null);
});

test('the highlight follows speech boundaries, then the chunks, and never writes to the document', () => {
    const {ctl, doc, spoken} = loadHl([['Hello wor', 'ld. Second one.'], ['Third part.']]);
    ctl.start();
    assert.equal(doc.calls.filter(c => c === 'GetSelectionState').length, 1);
    // before the voice reports anything the whole chunk is shown
    assert.equal(selectedText(doc), 'Hello world. Second one.');
    spoken[0].onstart();
    spoken[0].onboundary({name: 'word', charIndex: 0});
    assert.equal(selectedText(doc), 'Hello world.');
    spoken[0].onboundary({name: 'word', charIndex: 6});
    assert.equal(selectedText(doc), 'Hello world.');
    spoken[0].onboundary({name: 'word', charIndex: 13});
    assert.equal(selectedText(doc), 'Second one.');
    spoken[0].onboundary({name: 'line', charIndex: 0});         // not a word or sentence boundary: ignored
    assert.equal(selectedText(doc), 'Second one.');
    spoken[0].onend();
    assert.equal(spoken[1].text, 'Third part.');
    assert.equal(selectedText(doc), 'Third part.');
    ctl.stop();
    assert.equal(doc.calls[doc.calls.length - 1] === 'ruler' && doc.calls.includes('SetSelectionState'), true);
    assert.equal(doc.selection, doc.state);                       // the selection of the user is back
    assert.equal(doc.calls.filter(c => c === 'GetSelectionState').length, 1);
});

test('a voice without boundary events keeps the chunk shown, and late events of old utterances are ignored', () => {
    const {ctl, doc, spoken} = loadHl([['One. Two.'], ['Three.']]);
    ctl.start();
    spoken[0].onstart();
    assert.equal(selectedText(doc), 'One. Two.');
    ctl.pause();
    const sel = doc.selection;
    spoken[0].onboundary({name: 'word', charIndex: 5});
    assert.equal(doc.selection, sel);
    ctl.resume();
    spoken[1].onstart();
    assert.equal(selectedText(doc), 'One. Two.');
    ctl.stop();
});

test('a click in the document gives up the saved selection; typing stops the reading and restores it', () => {
    const a = loadHl([['One. Two.']]);
    a.ctl.start();
    a.ctl._saved = null;                                         // what the mousedown handler does
    a.ctl.stop();
    assert.ok(!a.doc.calls.includes('SetSelectionState'));

    const b = loadHl([['One. Two.']]);
    b.ctl.start();
    b.ctl.onKey({keyCode: 65, key: 'a', ctrlKey: false, altKey: false, shiftKey: false, target: {}, preventDefault() {}, stopPropagation() {}});
    assert.equal(b.ctl._state, 'idle');
    assert.equal(b.doc.selection, b.doc.state);
    // arrows and Ctrl+C keep reading
    const c = loadHl([['One. Two.']]);
    c.ctl.start();
    for (const k of [{keyCode: 40, key: 'ArrowDown'}, {keyCode: 67, key: 'c', ctrlKey: true}, {keyCode: 16, key: 'Shift'}]) {
        c.ctl.onKey(Object.assign({altKey: false, shiftKey: false, ctrlKey: false, target: {}, preventDefault() {}, stopPropagation() {}}, k));
    }
    assert.equal(c.ctl._state, 'playing');
});

test('content the highlight cannot place is read without highlight and without errors', () => {
    const {ctl, doc, spoken} = loadHl([['Hello.']]);
    doc.paras[0].GetText = () => 'completely different';
    ctl._paras = [{p: doc.paras[0], t: 'completely different'}];
    ctl.start();
    spoken[0].onstart();
    spoken[0].onboundary({name: 'word', charIndex: 0});
    assert.ok(!doc.calls.includes('select'));
    assert.equal(ctl._state, 'playing');
    ctl.stop();
    assert.ok(!doc.calls.includes('SetSelectionState'));         // nothing was selected, nothing to give back
});

test('isEditingKey tells editing keys from navigation', () => {
    const {C} = load();
    const k = (key, keyCode, extra) => Object.assign({key, keyCode, ctrlKey: false, altKey: false, metaKey: false}, extra);
    assert.equal(C.isEditingKey(k('a', 65)), true);
    assert.equal(C.isEditingKey(k('Backspace', 8)), true);
    assert.equal(C.isEditingKey(k('Enter', 13)), true);
    assert.equal(C.isEditingKey(k('v', 86, {ctrlKey: true})), true);
    assert.equal(C.isEditingKey(k('c', 67, {ctrlKey: true})), false);
    assert.equal(C.isEditingKey(k('ArrowLeft', 37)), false);
    assert.equal(C.isEditingKey(k('PageDown', 34)), false);
    assert.equal(C.isEditingKey(k('F5', 116)), false);
});

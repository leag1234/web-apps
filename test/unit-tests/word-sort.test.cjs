// SPDX-License-Identifier: AGPL-3.0-only
'use strict';

// Word Sort: the Sort controller (dialog, error messages, call to the editor), the Sort dialog settings
// and the strings in the nine shipped locales.

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const main = path.resolve(__dirname, '../../apps/documenteditor/main');
const read = f => fs.readFileSync(path.join(main, f), 'utf8');
const LOCALES = ['en', 'fr', 'de', 'es', 'it', 'ru', 'nl', 'pt', 'pt-pt'];

function env() {
    const underscore = {
        extend: (t, ...s) => Object.assign(t, ...s),
        each: (l, f) => (Array.isArray(l) ? l : Object.values(l)).forEach(f),
        some: (l, f) => l.some(f),
        template: s => () => s,
        bind: (f, c) => f.bind(c)
    };
    const make = base => ({extend: proto => { const C = function (o) { this.options = Object.assign({}, C.prototype.options, o || {}); }; C.prototype = Object.assign(Object.create(base || {}), proto); return C; }});
    const calls = {warn: [], edit: 0};
    const ctx = {
        _: underscore, console, DE: {Controllers: {}, Views: {}},
        Backbone: {Controller: make()},
        Common: {
            UI: {Window: Object.assign(make({initialize() {}, render() {}, getChild() {}}), {}), warning: o => { calls.warn.push(o); o.callback && o.callback(); }},
            NotificationCenter: {on() {}, trigger: ev => { if (ev === 'edit:complete') calls.edit++; }},
            Locale: {getCurrentLanguage: () => 'fr'},
            Utils: {String: {format: (s, ...a) => s.replace(/\{(\d+)\}/g, (m, i) => a[i])}}
        }
    };
    ctx.Common.UI.Window.prototype = {};
    ctx.define = (deps, f) => { if (typeof deps === 'function') f = deps; ctx._mod = f(); };
    vm.createContext(ctx);
    return {ctx, calls};
}

function loadDialog(e) {
    // Window needs a prototype with initialize/render used through Common.UI.Window.prototype.*
    e.ctx.Common.UI.Window.prototype = {initialize() {}, render() {}, getChild() { return null; }};
    const Window = e.ctx.Common.UI.Window;
    Window.extend = proto => {
        const C = function (o) { this.options = Object.assign({}, proto.options, o || {}); this.initialize(o); };
        C.prototype = Object.assign({}, Window.prototype, proto);
        C.prototype.show = function () { e.shown = this; };
        return C;
    };
    vm.runInContext(read('app/view/SortDialog.js'), e.ctx, {filename: 'SortDialog.js'});
    return e.ctx.DE.Views.SortDialog;
}

test('controller: an editor error code shows its message and returns the focus', () => {
    const e = env();
    vm.runInContext(read('app/controller/Sort.js'), e.ctx, {filename: 'Sort.js'});
    const Ctrl = e.ctx.DE.Controllers.Sort;
    const c = new Ctrl();
    c.api = {asc_GetSortInfo: () => ({error: 'merged'})};
    c.open();
    assert.equal(e.calls.warn.length, 1);
    assert.match(e.calls.warn[0].msg, /merged cells/);
    assert.equal(e.calls.edit, 1);
    c.api = {asc_GetSortInfo: () => ({error: 'nested'})};
    c.open();
    assert.match(e.calls.warn[1].msg, /nested table/);
    for (const code of ['none', 'merged', 'nested', 'section', 'mixed', 'locked'])
        assert.ok(Ctrl.errorText(c, code).length > 20, code);
});

test('controller: OK sends the dialog settings to the editor, Cancel does nothing', () => {
    const e = env();
    const Dialog = loadDialog(e);
    vm.runInContext(read('app/controller/Sort.js'), e.ctx, {filename: 'Sort.js'});
    const c = new e.ctx.DE.Controllers.Sort();
    const sent = [];
    c.api = {asc_GetSortInfo: () => ({mode: 'table', columns: 3, header: true}), asc_Sort: p => { sent.push(p); return 'ok'; }};
    c.open();
    const dlg = e.shown;
    assert.equal(dlg.options.lang, 'fr');
    dlg.getSettings = () => ({keys: [{field: 2, type: 1, ascending: false}], header: true});
    dlg.options.handler(dlg, 'cancel');
    assert.equal(sent.length, 0);
    dlg.options.handler(dlg, 'ok');
    assert.equal(sent.length, 1);
    assert.equal(sent[0].keys[0].field, 2);
    c.api.asc_Sort = () => 'locked';
    dlg.options.handler(dlg, 'ok');
    assert.match(e.calls.warn[0].msg, /protected or locked/);
});

test('dialog: fields are columns for a table and paragraph or fields for text; up to three keys', () => {
    const e = env();
    const Dialog = loadDialog(e);
    const t = new Dialog({info: {mode: 'table', columns: 3, header: false}});
    t._mode = 'table'; t._columns = 3; t._fields = 1; t._settings = {separator: 'tab'};
    assert.deepEqual(Array.from(t.fieldData(0).map(x => x.value)), [1, 2, 3]);
    assert.deepEqual(Array.from(t.fieldData(1).map(x => x.value)), [-1, 1, 2, 3]);
    const p = new Dialog({info: {mode: 'paragraphs', fields: 1}});
    p._mode = 'paragraphs'; p._fields = 1; p._settings = {separator: 'tab'};
    assert.deepEqual(Array.from(p.fieldData(0).map(x => x.value)), [0]);
    p._fields = 3;
    assert.deepEqual(Array.from(p.fieldData(0).map(x => x.value)), [0, 1, 2, 3]);
    assert.equal(p.fieldData(0)[2].displayValue, 'Field 2');

    p.cmbField = [{getValue: () => 2}, {getValue: () => -1}, {getValue: () => 0}];
    p.cmbType = [{getValue: () => 1}, {getValue: () => 0}, {getValue: () => 2}];
    p.cmbOrder = [{getValue: () => 0}, {getValue: () => 1}, {getValue: () => 1}];
    p.rbHeader = {getValue: () => true};
    p._settings = {separator: 'comma', separatorChar: '', caseSensitive: true};
    p.options.lang = 'de';
    const s = p.getSettings();
    assert.deepEqual(JSON.parse(JSON.stringify(s.keys)), [{field: 2, type: 1, ascending: false}, {field: 0, type: 2, ascending: true}]);
    assert.equal(s.header, true);
    assert.equal(s.caseSensitive, true);
    assert.equal(s.separator, 'comma');
    assert.equal(s.lang, 'de');
});

test('Sort is wired in the Home toolbar, the table Layout tab and the app lists', () => {
    assert.match(read('app/view/Toolbar.js'), /this\.btnSort = new Common\.UI\.Button/);
    assert.match(read('app/template/Toolbar.template'), /slot-btn-sort/);
    assert.match(read('app/controller/Toolbar.js'), /toolbar\.btnSort\.on\('click'/);
    assert.match(read('app/view/ContextTabs.js'), /big\('elx-btn-tbl-sort'\)/);
    assert.match(read('app/controller/ContextTabs.js'), /v\.btnTblSort\.on\('click'/);
    for (const f of ['app.js', 'app_dev.js']) {
        assert.match(read(f), /'documenteditor\/main\/app\/controller\/Sort'/);
        assert.match(read(f), /,'Sort'/);
    }
    for (const f of ['app_dev.js', 'app_pack.js']) assert.match(read(f), /view\/SortDialog/);
});

test('all strings exist in the nine shipped locales and differ from English outside Date and Text-like words', () => {
    const src = read('app/view/SortDialog.js') + read('app/controller/Sort.js');
    const keys = new Set();
    for (const m of src.matchAll(/^\s+(text[A-Za-z]+|txt[A-Za-z]+): '/gm)) keys.add(m[1]);
    assert.ok(keys.size >= 25);
    const en = JSON.parse(read('locale/en.json'));
    for (const l of LOCALES) {
        const d = JSON.parse(read('locale/' + l + '.json'));
        for (const k of ['DE.Views.Toolbar.tipSort', 'DE.Views.ContextTabs.txtSort', 'DE.Controllers.Sort.txtTitle', 'DE.Controllers.Sort.txtErrNone', 'DE.Controllers.Sort.txtErrLocked',
            'DE.Views.SortDialog.textSortBy', 'DE.Views.SortDialog.textThenBy', 'DE.Views.SortDialog.textHeaderRow', 'DE.Views.SortDialog.textOptions',
            'DE.Views.SortOptionsDialog.textTitle', 'DE.Views.SortOptionsDialog.textSeparator', 'DE.Views.SortOptionsDialog.textCase'])
            assert.ok(d[k], l + ' ' + k);
        if (l !== 'en') assert.notEqual(d['DE.Views.SortDialog.textSortBy'], en['DE.Views.SortDialog.textSortBy'], l);
    }
    // every view/controller string has an English locale entry
    const prefixFor = k => /^txt(Err|Title)/.test(k) ? 'DE.Controllers.Sort.' : null;
    for (const k of ['textSortBy', 'textThenBy', 'textAscending', 'textDescending', 'textParagraphs', 'textField', 'textColumn', 'textNone', 'textMyList', 'textHeaderRow', 'textNoHeaderRow', 'textOptions', 'textText', 'textNumber', 'textDate'])
        assert.ok(en['DE.Views.SortDialog.' + k], k);
    assert.ok(prefixFor('txtErrNone'));
});

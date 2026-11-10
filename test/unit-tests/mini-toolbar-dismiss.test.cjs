// SPDX-License-Identifier: AGPL-3.0-only
'use strict';
// Mini Toolbar on right-click: closes with the context menu, on any dismissing action, and honours the setting.
const test = require('node:test'), assert = require('node:assert/strict'), fs = require('node:fs'), path = require('node:path'), vm = require('node:vm');
const root = path.resolve(__dirname, '../..');

function load(file, key, ns) {
    const store = {}, docHandlers = {}, docListeners = {}, callbacks = {};
    const $ = selector => ({on: (name, fn) => { docHandlers[name] = fn; }, is: () => false, find: () => $('x'), addBack: () => $('x'), first: () => $('x'), length: 1,
        offset: () => ({left: 50, top: 300}), outerHeight: () => 200, outerWidth: () => 150, cmpEl: null});
    $.fn = {};
    const _ = {extend: Object.assign, bind: (f, t) => f.bind(t), defer: f => f(), each: (l, f) => (l || []).forEach(f)};
    const context = {console, setTimeout: f => f(), clearTimeout() {}, _, $, document: {addEventListener: (n, f) => { docListeners[n] = f; }},
        [ns]: {Controllers: {}}, Backbone: {Controller: {extend: spec => spec}},
        Common: {NotificationCenter: {on() {}}, Utils: {ModalWindow: {isVisible: () => false}, innerWidth: () => 1200, String: {htmlEncode: s => s}},
            localStorage: {getBool: (k, d) => k in store ? !!store[k] : d}, UI: {Menu: {Manager: {hideAll() {}}}}},
        define(_deps, factory) { factory(); }};
    vm.runInNewContext(fs.readFileSync(path.join(root, file), 'utf8'), context);
    const ctrl = context[ns].Controllers[key];
    return {ctrl, store, docHandlers, docListeners, callbacks};
}
function menuFixture(ctrl, visibleClass) {
    const hidden = [], handlers = {};
    const menu = {isVisible: () => true, cmpEl: {find: () => ({addBack: () => ({first: () => ({length: 1, offset: () => ({left: 50, top: 300}), outerHeight: () => 200, outerWidth: () => 150})})})},
        once: (name, fn) => { handlers[name] = fn; }};
    const holder = {ssMenu: menu, textMenu: menu, currentMenu: menu};
    ctrl.getApplication = () => ({getController: name => name === 'DocumentHolder' ? {documentHolder: holder, currentMenu: menu} : {}});
    return {menu, handlers, hidden};
}
function excel(opts = {}) {
    const e = load('apps/spreadsheeteditor/main/app/controller/ExcelMiniToolbar.js', 'ExcelMiniToolbar', 'SSE'), c = e.ctrl;
    const classes = new Set();
    c.$el = {addClass(n) { classes.add(n); return this; }, removeClass(n) { classes.delete(n); return this; }, hasClass: n => classes.has(n), css() { return this; },
        outerWidth: () => 200, outerHeight: () => 60, 0: {contains: () => false}};
    c.mode = {isEdit: true}; c.tb = {btnCopyStyle: null}; c.syncState = () => {}; c.isOwnMenu = () => false; c.isMenuOpen = () => false;
    const m = menuFixture(c); return {...e, c, classes, ...m};
}

test('Excel: the mini toolbar appears with the context menu and closes when the menu closes after a command', () => {
    const e = excel(); e.c.onContextMenu();
    assert(e.classes.has('visible')); assert(e.handlers['hide:after'], 'closes with the menu');
    e.handlers['hide:after'](); assert(!e.classes.has('visible'));
});
test('Excel: using the mini toolbar keeps it for more formatting; a new right-click starts fresh', () => {
    const e = excel(); e.c.onContextMenu(); e.c.interacting = true; e.handlers['hide:after'](); assert(e.classes.has('visible'));
    e.c.onContextMenu(); assert.equal(e.c.interacting, false, 'stale interaction reset');
    e.handlers['hide:after'](); assert(!e.classes.has('visible'));
});
test('Excel: the right-click setting turns it off', () => {
    const e = excel(); e.store['eo-minitoolbar-context'] = 0; e.c.onContextMenu(); assert(!e.classes.has('visible'));
    e.store['eo-minitoolbar-context'] = 1; e.c.onContextMenu(); assert(e.classes.has('visible'));
});
test('Excel: selection change, sheet switch, scrolling, typing and clicking elsewhere dismiss it', () => {
    const cases = ['selection', 'sheet', 'wheel', 'keydown', 'mousedown'];
    for (const how of cases) {
        const e = excel(), c = e.c, api = {asc_registerCallback: (n, f) => { e.callbacks[n] = f; }};
        c.getApplication = () => ({getController: n => n === 'Viewport' ? {getApi: () => api} : n === 'Toolbar' ? {toolbar: {}} : {}});
        c.render = () => {}; c.onAppReady({isEdit: true});
        menuFixture(c); c.$el = excel().c.$el; const classes = new Set(['visible']);
        c.$el = {addClass() {}, removeClass(n) { classes.delete(n); }, hasClass: n => classes.has(n), 0: {contains: () => false}};
        c.isOwnMenu = () => false; c.isMenuOpen = () => false; c.interacting = false;
        if (how === 'selection') e.callbacks.asc_onSelectionChanged();
        else if (how === 'sheet') e.callbacks.asc_onActiveSheetChanged();
        else if (how === 'wheel') e.docListeners.wheel({target: {}});
        else e.docHandlers[how]({target: {}});
        assert(!classes.has('visible'), how);
    }
});
test('Excel: its own formatting does not dismiss it through a selection update', () => {
    const e = excel(), c = e.c, api = {asc_registerCallback: (n, f) => { e.callbacks[n] = f; }};
    c.getApplication = () => ({getController: n => n === 'Viewport' ? {getApi: () => api} : n === 'Toolbar' ? {toolbar: {}} : {}});
    c.render = () => {}; c.onAppReady({isEdit: true});
    const classes = new Set(['visible']); c.$el = {removeClass(n) { classes.delete(n); }, hasClass: n => classes.has(n), 0: {contains: () => false}};
    c.interacting = true; e.callbacks.asc_onSelectionChanged(); assert(classes.has('visible'));
});
test('Word: right-click honours the setting and resets a stale interaction', () => {
    const e = load('apps/documenteditor/main/app/controller/MiniToolbar.js', 'MiniToolbar', 'DE'), c = e.ctrl, shown = [];
    const m = menuFixture(c); c.view = {setTableMode() {}}; c.canShow = () => true; c.showAbove = () => shown.push(1);
    e.store['eo-minitoolbar-context'] = 0; c.interacting = true; c.onContextMenu(); assert.equal(shown.length, 0); assert.equal(c.interacting, false);
    e.store['eo-minitoolbar-context'] = 1; c.onContextMenu(); assert.equal(shown.length, 1); assert(m.handlers['hide:after']);
    let hid = 0; c.hide = () => hid++; m.handlers['hide:after'](); assert.equal(hid, 1, 'closes with the context menu');
});

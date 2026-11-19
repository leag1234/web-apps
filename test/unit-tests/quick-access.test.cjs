// SPDX-License-Identifier: AGPL-3.0-only
'use strict';
// "Show Quick Access Toolbar": hides the quick access buttons, shrinks the title row, persists, reports changes to the host.
const test = require('node:test'), assert = require('node:assert/strict'), fs = require('node:fs'), path = require('node:path'), vm = require('node:vm');
const root = path.resolve(__dirname, '../..');

function load({stored, compact = false, titleHeight = 36, visibleItem = true} = {}) {
    const store = stored === undefined ? {} : {'elx-show-quick-access': stored ? 1 : 0};
    const bodyClasses = new Set(compact ? ['eo-compact-editor'] : []);
    const body = {toggleClass(n, on) { on ? bodyClasses.add(n) : bodyClasses.delete(n); return this; }, hasClass: n => bodyClasses.has(n)};
    const item = {height: titleHeight, el: {is: () => visibleItem}};
    const calls = {doLayout: 0, triggers: [], events: [], menus: [], prevented: 0, stopped: 0, settings: {}};
    const layout = {getItem: name => name === 'title' ? item : null, doLayout() { calls.doLayout++; }};
    const nodeOf = text => { const n = {text: '', click: null, on(name, fn) { if (name === 'click') n.click = fn; return n; }, text_(v) { n.text = v; return n; }, focus() {}}; return n; };
    const $ = arg => {
        if (arg === 'body') return body;
        if (arg === '#toolbar') return {length: 1, offset: () => ({top: 36})};
        if (arg && arg.__target) return {closest: sel => ({length: arg.matches.includes(sel) ? 1 : 0})};
        if (typeof arg === 'string' && arg.startsWith('<div')) {
            const button = {text(v) { calls.menus.push(v); return button; }, on() { return button; }, focus() {}};
            const menu = {appendTo() { return menu; }, find: () => button, outerWidth: () => 220, outerHeight: () => 30, css: () => menu, remove() { calls.menus.push('removed'); }};
            return menu;
        }
        return {on() { return this; }, off() { return this; }, scrollTop: () => 0, length: 0, closest: () => ({length: 0})};
    };
    const handlers = {}, notify = {};
    const window = {DE: {getController: () => ({viewport: {vlayout: layout}, header: {txtHideQuickAccess: 'Hide it', txtShowQuickAccess: 'Show it'}})},
        innerWidth: 1360, innerHeight: 900, dispatchEvent(e) { calls.events.push(e.detail); return true; }, on() {}};
    const context = {console, setTimeout: f => f(), $, window, document: {head: {appendChild() {}}, createElement: () => ({}), addEventListener: (n, f) => { handlers[n] = f; }, body: {}},
        CustomEvent: class { constructor(type, init) { this.type = type; this.detail = init.detail; } },
        Common: {UI: {Menu: {Manager: {hideAll() {}}}}, localStorage: {getBool: (k, d) => k in store ? !!store[k] : d, setBool(k, v) { store[k] = v ? 1 : 0; }},
            Utils: {InternalSettings: {set: (k, v) => { calls.settings[k] = v; }, get: k => k === 'toolbar-height-tabs' ? 32 : undefined}},
            NotificationCenter: {on: (n, f) => { notify[n] = f; }, trigger: (...a) => calls.triggers.push(a.join(':'))}},
        define(_deps, factory) { return factory(); }};
    context.window.window = context.window;
    vm.runInNewContext(fs.readFileSync(path.join(root, 'apps/common/main/lib/elixir-chrome/QuickAccess.js'), 'utf8'), context);
    return {api: context.window.ElixirQuickAccess, store, bodyClasses, item, calls, handlers, notify};
}

test('the quick access toolbar is shown by default and the row keeps its height', () => {
    const q = load(); q.notify['app:ready']();
    assert.equal(q.api.isVisible(), true); assert(!q.bodyClasses.has('elx-no-qat')); assert.equal(q.item.height, 36);
});
test('hiding adds the class at once, shrinks the title row to 28px, relays out, saves and tells the host once', () => {
    const q = load(); q.notify['app:ready']();
    q.api.set(false);
    assert(q.bodyClasses.has('elx-no-qat')); assert.equal(q.item.height, 28); assert.equal(q.calls.settings['document-title-height'], 28);
    assert.equal(q.calls.doLayout, 1); assert(q.calls.triggers.includes('layout:changed:header')); assert.equal(q.store['elx-show-quick-access'], 0);
    assert.equal(JSON.stringify(q.calls.events), JSON.stringify([{visible: false}]));
    q.api.set(false); assert.equal(q.calls.events.length, 1, 'no event without a change');
});
test('showing again restores the original height', () => {
    const q = load(); q.notify['app:ready'](); q.api.set(false); q.api.set(true);
    assert(!q.bodyClasses.has('elx-no-qat')); assert.equal(q.item.height, 36); assert.equal(q.calls.settings['document-title-height'], 36);
    assert.equal(q.store['elx-show-quick-access'], 1); assert.equal(JSON.stringify(q.calls.events), JSON.stringify([{visible: false}, {visible: true}]));
});
test('a stored "hidden" choice is applied before the editor is ready (no flash) and the layout when it is', () => {
    const q = load({stored: false});
    assert(q.bodyClasses.has('elx-no-qat')); assert.equal(q.item.height, 36);
    q.notify['app:ready'](); assert.equal(q.item.height, 28);
});
test('a change coming from the host is applied and saved but not reported back', () => {
    const q = load(); q.notify['app:ready'](); q.api.set(false, {host: true});
    assert(q.bodyClasses.has('elx-no-qat')); assert.equal(q.item.height, 28); assert.equal(q.calls.events.length, 0);
});
test('a title row that is already smaller than 28px is not made taller, and a hidden or empty title row is left alone', () => {
    const small = load({titleHeight: 24}); small.notify['app:ready'](); small.api.set(false); assert.equal(small.item.height, 24);
    const zero = load({titleHeight: 0}); zero.notify['app:ready'](); zero.api.set(false); assert.equal(zero.calls.doLayout, 0);
    const hiddenItem = load({visibleItem: false}); hiddenItem.notify['app:ready'](); hiddenItem.api.set(false); assert.equal(hiddenItem.calls.doLayout, 0);
});
test('with the compact desktop title bar the row height is not touched; refresh() restores it afterwards', () => {
    const q = load({compact: true}); q.notify['app:ready'](); q.api.set(false); assert.equal(q.item.height, 36); assert.equal(q.calls.doLayout, 0);
    q.bodyClasses.delete('eo-compact-editor'); q.api.refresh(); assert.equal(q.item.height, 28);
});
test('right-click on the quick access toolbar or the ribbon tab row opens the menu; elsewhere nothing happens', () => {
    const q = load(); q.notify['app:ready']();
    const ev = (matches, y = 10) => { const e = {__target: true, target: null, clientX: 300, clientY: y, preventDefault() { q.calls.prevented++; }, stopPropagation() { q.calls.stopped++; }}; e.target = e; e.matches = matches; return e; };
    q.handlers.contextmenu(ev(['#box-document-title > .hedset[role="menubar"]']));
    assert.equal(q.calls.prevented, 1); assert(q.calls.menus.includes('Hide it'));
    q.handlers.contextmenu(ev(['#toolbar'], 50));      // tab row: 36..68
    assert.equal(q.calls.prevented, 2);
    q.handlers.contextmenu(ev(['#toolbar'], 100));     // ribbon commands: not ours
    assert.equal(q.calls.prevented, 2);
    q.handlers.contextmenu(ev(['#toolbar', '.btn, button, input, select, textarea, .dropdown-menu, .combobox, .form-control'], 50)); // a button on the tab row
    assert.equal(q.calls.prevented, 2);
    q.handlers.contextmenu(ev([]));
    assert.equal(q.calls.prevented, 2);
});
test('the menu says "Show" while the toolbar is hidden', () => {
    const q = load(); q.notify['app:ready'](); q.api.set(false);
    const e = {__target: true, clientX: 10, clientY: 50, preventDefault() {}, stopPropagation() {}, matches: ['#toolbar']}; e.target = e;
    q.handlers.contextmenu(e); assert(q.calls.menus.includes('Show it'));
});
test('the compact desktop title bar keeps the native context menu (its quick buttons follow the setting in the shell)', () => {
    const q = load({compact: true}); q.notify['app:ready']();
    const e = {__target: true, clientX: 10, clientY: 50, preventDefault() { q.calls.prevented++; }, stopPropagation() {}, matches: ['#toolbar']}; e.target = e;
    q.handlers.contextmenu(e); assert.equal(q.calls.prevented, 0);
});

const EDITORS = ['documenteditor', 'spreadsheeteditor', 'presentationeditor', 'pdfeditor'];
const LOCALES = ['en', 'fr', 'de', 'es', 'it', 'pt', 'pt-pt', 'nl', 'pl', 'ru', 'ja', 'zh'];
test('both strings exist in the header, which loads the module, and in every editor locale that carries Elixir strings', () => {
    const header = fs.readFileSync(path.join(root, 'apps/common/main/lib/view/Header.js'), 'utf8');
    assert(header.includes("'common/main/lib/elixir-chrome/QuickAccess'"));
    assert(header.includes("txtHideQuickAccess: 'Hide Quick Access Toolbar'") && header.includes("txtShowQuickAccess: 'Show Quick Access Toolbar'"));
    for (const editor of EDITORS) for (const locale of LOCALES) {
        const file = path.join(root, 'apps', editor, 'main/locale', locale + '.json');
        if (!fs.existsSync(file)) continue;
        const json = JSON.parse(fs.readFileSync(file, 'utf8'));
        for (const key of ['Common.Views.Header.txtHideQuickAccess', 'Common.Views.Header.txtShowQuickAccess']) {
            assert.equal(typeof json[key], 'string', editor + ' ' + locale + ' ' + key); assert(json[key].trim());
        }
        if (locale === 'en') {
            assert.equal(json['Common.Views.Header.txtHideQuickAccess'], 'Hide Quick Access Toolbar');
            assert.equal(json['Common.Views.Header.txtShowQuickAccess'], 'Show Quick Access Toolbar');
        } else assert.notEqual(json['Common.Views.Header.txtHideQuickAccess'], 'Hide Quick Access Toolbar', editor + ' ' + locale);
    }
});

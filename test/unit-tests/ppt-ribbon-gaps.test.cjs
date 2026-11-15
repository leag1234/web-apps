// SPDX-FileCopyrightText: 2026 ElixirOffice contributors
// SPDX-License-Identifier: AGPL-3.0-only
'use strict';

// PowerPoint table ribbon (borders, cell size, margins, text direction), Design > Fonts and Background
// Styles, and the shared WordArt Styles group (Word and PowerPoint).

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const root = path.resolve(__dirname, '../..');
const apps = path.join(root, 'apps');
const read = f => fs.readFileSync(path.join(apps, f), 'utf8');
const plain = v => JSON.parse(JSON.stringify(v));

// ---- a tiny mock of the Asc property classes: every put_/asc_put call is recorded ----------------
function ascMock() {
    const Asc = {c_oAscFill: {FILL_TYPE_NOFILL: 0, FILL_TYPE_SOLID: 1, FILL_TYPE_GRAD: 3}, c_oAscStrokeType: {STROKE_NONE: 0, STROKE_COLOR: 1},
        c_oAscFillGradType: {GRAD_LINEAR: 1}, c_oAscCellTextDirection: {LRTB: 0, TBRL: 1, BTLR: 2}};
    const rec = name => {
        Asc[name] = function () { this.v = {}; };
        Asc[name].prototype = new Proxy({}, {get: (t, k) => {
            if (typeof k !== 'string') return undefined;
            if (/^(put_|asc_put)/.test(k)) return function (val) { this.v[k.replace(/^(put_|asc_put)/, '')] = val; };
            if (/^(get_|asc_get)/.test(k)) return function () { return this.v[k.replace(/^(get_|asc_get)/, '')]; };
            return undefined;
        }});
    };
    ['asc_TextArtProperties', 'asc_CShapeFill', 'asc_CFillSolid', 'asc_CStroke', 'asc_CShapeProperty', 'asc_CTextBorder', 'CBorders',
        'CTableProp', 'CMargins', 'asc_CColor', 'CAscSlideProps', 'asc_CFillGrad'].forEach(rec);
    return Asc;
}

function ui() {
    const handlers = new WeakMap();
    class Emitter {
        constructor(o) { this.options = o || {}; Object.assign(this, o || {}); handlers.set(this, {}); }
        on(ev, fn) { (handlers.get(this)[ev] = handlers.get(this)[ev] || []).push(fn); }
        trigger(ev, ...a) { (handlers.get(this)[ev] || []).forEach(f => f(...a)); }
        render() { this.rendered = true; }
        setMenu() { this.menu = this.menu || {}; }
        getPicker() { return this.picker = this.picker || {updateColors() { this.updated = true; }}; }
        isAutoColor() { return this.color === 'auto'; }
        setColor(c) { this.color = c; }
    }
    class Menu extends Emitter { constructor(o) { super(o); this.items = (o.items || []); } setInnerMenu() {} hide() {} }
    class MenuItem extends Emitter {}
    class DataViewStore { constructor(m) { this.models = m; } }
    class DataView extends Emitter {}
    return {Emitter, Menu, MenuItem, DataView, DataViewStore, trigger: (o, ev, ...a) => o.trigger(ev, ...a)};
}

function loadModule(file, extra, ns) {
    const U = ui(), Asc = ascMock();
    const context = Object.assign({
        Asc, _: Object.assign(((x) => x), {
            extend: Object.assign, each: (a, f) => (Array.isArray(a) ? a.forEach(f) : Object.keys(a || {}).forEach(k => f(a[k], k))),
            map: (a, f) => a.map(f), bind: (f, o) => f.bind(o), template: s => () => s, isObject: o => o !== null && typeof o === 'object',
            keys: Object.keys
        }),
        $: () => ({}),
        Common: {
            UI: {Button: U.Emitter, ButtonColored: U.Emitter, Menu: U.Menu, MenuItem: U.MenuItem, DataView: U.DataView, DataViewStore: U.DataViewStore},
            Utils: {ThemeColor: {getRgbColor: c => ({rgb: c}), getEffectColors: () => [{color: 'ffffff', effectId: 1}, {color: '000000', effectId: 2}, {color: 'eeeeee', effectId: 3},
                {color: '112233', effectId: 4}, {color: 'ff0000', effectId: 5}, {color: '00ff00', effectId: 6}]}, lockControls() {}},
            NotificationCenter: {trigger() {}, on() {}},
            Controllers: {}
        },
        Backbone: {
            Controller: {extend: props => { const C = function () {}; Object.assign(C.prototype, props); return C; }},
            View: {extend: props => { const C = function () {}; Object.assign(C.prototype, props); return C; }}
        },
        console
    }, extra);
    context[ns] = {Controllers: {}, Views: {}, getCollection: () => ({length: 1}), getController: () => ({})};
    context.window = context;
    context.define = (_deps, factory) => factory();
    vm.runInNewContext(fs.readFileSync(path.join(apps, file), 'utf8'), context, {filename: path.basename(file)});
    return context;
}

// ---- WordArt Styles (common) ------------------------------------------------------------------------
function wordArt() {
    const ctx = loadModule('common/main/lib/component/WordArtRibbon.js', {}, 'PE');
    const applied = [], added = {};
    const host = {
        strings: {txtWordArtQuickStyles: 'Quick Styles', txtTextFill: 'Text Fill', txtTextOutline: 'Text Outline', txtTextEffects: 'Text Effects', txtTransform: 'Transform', txtWeight: 'Weight'},
        ids: {styles: 'a', fill: 'b', outline: 'c', effects: 'd'},
        lock: ['x'],
        add: (id, cmp) => { added[id] = cmp; },
        apply: fn => { const s = new ctx.Asc.asc_CShapeProperty(); fn(s); applied.push(s.v.TextArtProperties.v); },
        api: () => ({asc_getPropertyEditorTextArts: () => [{Image: 'i.png', Type: 7}]}),
        collection: () => ({length: 1}),
        fillTextArt() {},
        done() { host.doneCount = (host.doneCount || 0) + 1; }
    };
    const wa = ctx.Common.UI.WordArtRibbon.create(host);
    return {ctx, wa, host, applied, added};
}

test('WordArt group registers its four controls in the slots the tab declares', () => {
    const {added, wa} = wordArt();
    assert.deepEqual(Object.keys(added), ['a', 'b', 'c', 'd']);
    assert.equal(wa.controls.length, 4);
    assert.equal(wa.buttons[0], wa.btnFill);
    assert.equal(wa.buttons[1], wa.btnOutline);
});

test('Text Fill applies a solid text art fill, or none for transparent', () => {
    const {wa, applied} = wordArt();
    wa.btnFill.trigger('color:select', wa.btnFill, 'ff0000');
    assert.equal(applied.length, 1);
    const fill = applied[0].Fill;
    assert.equal(fill.v.type, 1);
    assert.equal(fill.v.fill.v.color.rgb, 'ff0000');
    wa.btnFill.trigger('color:select', wa.btnFill, 'transparent');
    assert.equal(applied[1].Fill.v.type, 0);
});

test('Text Outline applies colour and weight in millimetres, and "no outline"', () => {
    const {wa, applied} = wordArt();
    wa.btnOutline.trigger('color:select', wa.btnOutline, '0000ff');
    const stroke = applied[0].Line.v;
    assert.equal(stroke.type, 1);
    assert.ok(Math.abs(stroke.width - 25.4 / 72) < 1e-9, 'default weight is 1 pt');
    wa.weight.menu.trigger('item:click', wa.weight.menu, {value: 3});
    assert.ok(Math.abs(applied[1].Line.v.width - 3 * 25.4 / 72) < 1e-9);
    assert.equal(applied[1].Line.v.color.rgb, '0000ff', 'the weight keeps the chosen colour');
    wa.btnOutline.trigger('color:select', wa.btnOutline, 'transparent');
    assert.equal(applied[2].Line.v.type, 0);
});

test('the pane colour pickers are refreshed when the theme colours change', () => {
    const {wa} = wordArt();
    wa.updateColors([1], [2]);
    assert.equal(wa.btnFill.picker.updated, true);
    assert.equal(wa.btnOutline.picker.updated, true);
});

test('the Weight list is the one of Shape Outline', () => {
    const {ctx} = wordArt();
    assert.deepEqual(plain(ctx.Common.UI.WordArtRibbon.WEIGHTS), [0.25, 0.5, 0.75, 1, 1.5, 2.25, 3, 4.5, 6]);
});

// ---- PowerPoint Table Design and Layout ---------------------------------------------------------------
function tabs() {
    const ctx = loadModule('presentationeditor/main/app/elixir-slide/ContextTabs.js', {}, 'PE');
    const proto = ctx.PE.Controllers.SlideContextTabs.prototype;
    const calls = [];
    const self = Object.assign(Object.create(proto), {
        api: {tblApply: p => calls.push(p)}, done() {}, _borderSize: 1,
        tdPen: {isAutoColor: () => true, color: 'auto'},
        rgb: c => ({rgb: c})
    });
    return {ctx, proto, self, calls};
}

test('Borders presets set only the edges they name, with the pen weight in mm', () => {
    const {self, calls} = tabs();
    const edges = b => Object.keys(b.v).sort().join(',');
    self.applyCellBorders('lrtb');
    let p = calls[0].v;
    assert.equal(edges(p.CellBorders), 'Bottom,Left,Right,Top');
    assert.equal(p.CellBorders.v.Top.v.Value, 1);
    assert.ok(Math.abs(p.CellBorders.v.Top.v.Size - 25.4 / 72) < 1e-9);
    assert.equal(p.CellBorders.v.Top.v.Color.v.auto, true, 'auto pen colour');
    assert.equal(p.CellSelect, true);
    self.applyCellBorders('cm');
    assert.equal(edges(calls[1].v.CellBorders), 'InsideH,InsideV');
    self.applyCellBorders('lrtbcm');
    assert.equal(edges(calls[2].v.CellBorders), 'Bottom,InsideH,InsideV,Left,Right,Top');
});

test('No Border clears all six edges', () => {
    const {self, calls} = tabs();
    self.applyCellBorders('');
    const b = calls[0].v.CellBorders.v;
    assert.deepEqual(Object.keys(b).sort(), ['Bottom', 'InsideH', 'InsideV', 'Left', 'Right', 'Top']);
    Object.values(b).forEach(e => assert.equal(e.v.Value, 0));
});

test('a zero pen weight draws no border', () => {
    const {self, calls} = tabs();
    self._borderSize = 0;
    self.applyCellBorders('b');
    assert.equal(calls[0].v.CellBorders.v.Bottom.v.Value, 0);
});

test('a chosen pen colour is used instead of auto', () => {
    const {self, calls} = tabs();
    self.tdPen = {isAutoColor: () => false, color: 'ff8800'};
    self.applyCellBorders('t');
    assert.equal(calls[0].v.CellBorders.v.Top.v.Color.rgb, 'ff8800');
});

test('Cell Margins presets: Normal uses the table defaults, the others set the four margins in mm', () => {
    const {self, calls} = tabs();
    self.applyCellMargins('normal');
    assert.equal(calls[0].v.CellMargins.v.Flag, 0);
    self.applyCellMargins('none');
    let m = calls[1].v.CellMargins.v;
    assert.deepEqual([m.Left, m.Top, m.Right, m.Bottom, m.Flag], [0, 0, 0, 0, 2]);
    self.applyCellMargins('wide');
    m = calls[2].v.CellMargins.v;
    assert.deepEqual([m.Left, m.Right], [5.08, 5.08]);
    assert.equal(calls[2].v.CellSelect, true);
});

test('the Table tabs use the strings, groups and locale keys they declare', () => {
    const src = read('presentationeditor/main/app/elixir-slide/ContextTabs.js');
    for (const k of ['txtGroupDrawBorders', 'txtGroupCellSize', 'txtGroupWordArt']) assert.match(src, new RegExp(k));
    assert.match(src, /put_CellsTextDirection\(it\.options\.direction\)/);
    assert.match(src, /put_RowHeight\(mm\)/);
    assert.match(src, /put_ColumnWidth\(mm\)/);
    assert.match(src, /'common\/main\/lib\/component\/WordArtRibbon'/);
});

// ---- Design > Fonts and Background Styles ---------------------------------------------------------------
function design() {
    const ctx = loadModule('presentationeditor/main/app/elixir-slide/DesignGalleries.js', {}, 'PE');
    const proto = ctx.PE.Controllers.SlideDesignGalleries.prototype;
    const sets = [], fonts = [];
    const self = Object.assign(Object.create(proto), {
        tb: {}, strings: {txtStyle: 'Style'},
        api: {SetSlideProps: p => sets.push(p.v), asc_ChangeFontScheme: (a, b, c) => fonts.push([a, b, c]), asc_getThemeFonts: () => ['Calibri Light', 'Calibri']}
    });
    return {ctx, self, sets, fonts};
}

test('font schemes: the Office list, unique names, a heading and a body font each', () => {
    const {ctx} = design();
    const list = plain(ctx.PE.Views.SlideDesignGalleries.FONT_SCHEMES);
    assert.ok(list.length >= 15);
    assert.equal(new Set(list.map(f => f[0])).size, list.length);
    assert.deepEqual(list[0], ['Office', 'Calibri Light', 'Calibri']);
    list.forEach(f => assert.ok(f[1] && f[2], f[0]));
});

test('choosing a font scheme sends heading font, body font and name to the engine', () => {
    const {self, fonts} = design();
    self.applyFonts(['Georgia', 'Georgia', 'Georgia']);
    assert.deepEqual(fonts, [['Georgia', 'Georgia', 'Georgia']]);
});

test('the current theme fonts are checked in the menu', () => {
    const {self} = design();
    const items = [{value: ['Office', 'Calibri Light', 'Calibri']}, {value: ['Arial', 'Arial', 'Arial']}].map(i => Object.assign(i, {setChecked(v) { this.checked = v; }}));
    self.markCurrentFonts({items});
    assert.deepEqual(items.map(i => i.checked), [true, false]);
});

test('twelve background styles: four colours as solid, subtle and intense', () => {
    const {ctx} = design();
    const rows = plain(ctx.PE.Views.SlideDesignGalleries.BG_STYLES);
    assert.equal(rows.length, 3);
    rows.forEach(r => assert.equal(r.length, 4));
    rows[0].forEach(spec => assert.equal(spec.length, 1));
    rows[1].concat(rows[2]).forEach(spec => assert.equal(spec.length, 2));
});

test('a solid background style is set on the slide, then applied to all slides', () => {
    const {self, sets} = design();
    self.applyBackground([2]);
    assert.equal(sets.length, 2);
    assert.equal(sets[0].background.v.type, 1);
    assert.equal(sets[0].background.v.fill.v.color.rgb.color, 'eeeeee');
    assert.equal(sets[1].ApplyBackgroundToAll, true);
});

test('a gradient background style has two stops from the theme colours', () => {
    const {self, sets} = design();
    self.applyBackground([0, 4]);
    const grad = sets[0].background.v;
    assert.equal(grad.type, 3);
    assert.deepEqual(plain(grad.fill.v.positions), [0, 100000]);
    assert.equal(grad.fill.v.colors.length, 2);
    assert.equal(grad.fill.v.colors[1].rgb.color, 'ff0000');
});

// ---- engine: theme fonts API ---------------------------------------------------------------------------------
test('the slide engine exposes asc_ChangeFontScheme and asc_getThemeFonts', () => {
    const sdk = path.resolve(root, '../sdkjs/slide');
    if (!fs.existsSync(path.join(sdk, 'api.js'))) return;   // sdkjs next to web-apps only
    const api = fs.readFileSync(path.join(sdk, 'api.js'), 'utf8');
    assert.match(api, /\['asc_ChangeFontScheme'\]/);
    assert.match(api, /\['asc_getThemeFonts'\]/);
});

// ---- locales -------------------------------------------------------------------------------------------------
const LANGS = ['en', 'fr', 'de', 'es', 'it', 'ru', 'nl', 'pt', 'pt-pt'];
const keysOf = (src, re) => [...src.matchAll(re)].map(m => m[1]);

test('every new string exists in the 9 shipped locales (presentation editor)', () => {
    const src = read('presentationeditor/main/app/elixir-slide/ContextTabs.js');
    const block = src.slice(src.indexOf('txtGroupDrawBorders'), src.indexOf('};', src.indexOf('txtGroupDrawBorders')));
    const keys = keysOf(block, /\b(txt\w+):/g);
    assert.ok(keys.length >= 30, 'new string keys found');
    const dg = read('presentationeditor/main/app/elixir-slide/DesignGalleries.js');
    const dgBlock = dg.slice(dg.indexOf('var STRINGS'), dg.indexOf('};', dg.indexOf('var STRINGS')));
    const dgKeys = keysOf(dgBlock, /\b((?:txt|tip)\w+):/g);
    for (const l of LANGS) {
        const loc = JSON.parse(read('presentationeditor/main/locale/' + l + '.json'));
        for (const k of keys) assert.ok(loc['PE.Views.SlideContextTabs.' + k], l + ' ' + k);
        for (const k of dgKeys) assert.ok(loc['PE.Views.SlideDesignGalleries.' + k], l + ' ' + k);
    }
});

test('every new string exists in the 9 shipped locales (document editor)', () => {
    const view = read('documenteditor/main/app/view/ContextTabs.js');
    const keys = ['txtGroupWordart', 'txtWordArtQuickStyles', 'txtTextFill', 'txtTextOutline', 'txtTextEffects', 'txtTransform'];
    keys.forEach(k => assert.match(view, new RegExp(k + ':')));
    for (const l of LANGS) {
        const loc = JSON.parse(read('documenteditor/main/locale/' + l + '.json'));
        for (const k of keys) assert.ok(loc['DE.Views.ContextTabs.' + k], l + ' ' + k);
    }
});

test('translated strings differ from English where they should', () => {
    const en = JSON.parse(read('presentationeditor/main/locale/en.json'));
    const fr = JSON.parse(read('presentationeditor/main/locale/fr.json'));
    assert.equal(en['PE.Views.SlideContextTabs.txtTextFill'], 'Text Fill');
    assert.equal(fr['PE.Views.SlideContextTabs.txtTextFill'], 'Remplissage du texte');
    assert.equal(fr['PE.Views.SlideDesignGalleries.txtFonts'], 'Polices');
});

test('both editors load the shared WordArt group and the presentation editor the Design galleries', () => {
    for (const f of ['presentationeditor/main/app.js', 'presentationeditor/main/app_dev.js']) {
        const s = read(f);
        assert.match(s, /elixir-slide\/DesignGalleries/, f);
        assert.match(s, /'SlideDesignGalleries'/, f);
    }
    assert.match(read('documenteditor/main/app/view/ContextTabs.js'), /Common\.UI\.WordArtRibbon\.create/);
});

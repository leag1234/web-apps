// SPDX-FileCopyrightText: 2026 ElixirOffice contributors
// SPDX-License-Identifier: AGPL-3.0-only
'use strict';

// Chart Design ribbon (Word, Excel, PowerPoint): Add Chart Element, Quick Layout, Change Colors and
// Chart Styles of the shared Common.UI.ChartRibbon.

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const apps = path.resolve(__dirname, '../../apps');
const plain = v => JSON.parse(JSON.stringify(v));
const read = f => fs.readFileSync(path.join(apps, f), 'utf8');

const LOCALES = ['en', 'fr', 'de', 'es', 'it', 'ru', 'nl', 'pt', 'pt-pt'];

// ---- mocks of the UI components the module uses ----------------------------------------------------
function ui() {
    const handlers = new WeakMap();
    class Emitter {
        constructor(o) { this.options = o || {}; Object.assign(this, o || {}); handlers.set(this, {}); }
        on(ev, fn) { (handlers.get(this)[ev] = handlers.get(this)[ev] || []).push(fn); return this; }
        trigger(ev, ...a) { (handlers.get(this)[ev] || []).forEach(f => f(...a)); }
        render() { this.rendered = true; }
    }
    class MenuItem extends Emitter { setChecked(v) { this.checked = v; } setVisible(v) { this.visible = v; } }
    class Menu extends Emitter {
        constructor(o) { super(o); this.items = (o.items || []).map(i => new MenuItem(i)); }
        addItem(i) { this.items.push(new MenuItem(i)); }
        removeAll() { this.items = []; }
        setInnerMenu(v) { this.inner = v; }
        hide() { this.hidden = true; }
    }
    class DataViewStore {
        constructor(m) { this.models = (m || []).map(x => ({get: k => x[k], set: (k, v) => { x[k] = v; }, attrs: x})); }
        reset(m) { DataViewStore.call(this, m); }
        findWhere(q) { return this.models.find(r => Object.keys(q).every(k => r.get(k) === q[k])); }
    }
    DataViewStore.call = (self, m) => { self.models = (m || []).map(x => ({get: k => x[k], set: (k, v) => { x[k] = v; }, attrs: x})); };
    class DataView extends Emitter { selectRecord(r) { this.selected = r; } deselectAll() { this.selected = null; } }
    return {Emitter, Menu, MenuItem, DataView, DataViewStore};
}

function loadModule() {
    const U = ui();
    const Asc = {
        c_oAscChartTitleShowSettings: {none: 0, overlay: 1, noOverlay: 2},
        c_oAscChartLegendShowSettings: {none: 0, left: 1, top: 2, right: 3, bottom: 4},
        c_oAscChartDataLabelsPos: {none: 0, ctr: 1, outEnd: 7, bestFit: 9},
        c_oAscChartTypeSettings: {barNormal: 0, hBarNormal: 5, pie: 4, pie3d: 14, comboBarLine: 40, comboBarLineSecondary: 41, comboAreaBar: 42, comboCustom: 43}
    };
    const lodash = {
        map: (a, f) => (Array.isArray(a) ? a : Object.values(a || {})).map(f), each: (a, f) => (Array.isArray(a) ? a.forEach(f) : Object.keys(a || {}).forEach(k => f(a[k], k))),
        filter: (a, f) => a.filter(f), pluck: (a, k) => a.map(x => x[k]), extend: Object.assign, template: s => () => s, bind: (f, o) => f.bind(o)
    };
    const history = {calls: [], StartTransaction() { this.calls.push('start'); }, EndTransaction() { this.calls.push('end'); }, Create_NewPoint() { this.calls.push('point'); }};
    const context = {
        Asc, AscCommon: {History: history}, AscDFH: {historyitem_type_ChartSpace: 1}, _: lodash, $: () => ({}),
        Common: {UI: {Button: U.Emitter, Menu: U.Menu, DataView: U.DataView, DataViewStore: U.DataViewStore}}, console
    };
    context.window = context;
    context.define = (_deps, factory) => factory();
    vm.runInNewContext(read('common/main/lib/component/ChartRibbon.js'), context, {filename: 'ChartRibbon.js'});
    return {ctx: context, U, history};
}

// the asc_ChartSettings of the engine, recording the calls
function chartProps(over) {
    const calls = [], rec = n => (...a) => { calls.push([n].concat(a)); };
    const p = Object.assign({
        calls, getType: () => 0, getStyle: () => 3, getTitle: () => 2, getLegendPos: () => 4, getDataLabelsPos: () => 0,
        getHorAxesProps: () => [{getLabel: () => 0, getGridlines: () => 0}], getVertAxesProps: () => [{getLabel: () => 0, getGridlines: () => 1}],
        getChartColorsAccent: () => -1,
        setDisplayChartTitle: rec('title'), setDisplayLegend: rec('legend'), setDisplayDataLabels: rec('labels'),
        setDisplayGridlines: rec('grid'), setDisplayAxisTitles: rec('axisTitles'), setChartColors: rec('colors'), putStyle: rec('putStyle')
    }, over || {});
    return p;
}

function build(opts) {
    opts = opts || {};
    const {ctx, U, history} = loadModule();
    const added = {}, applied = [], dhCalls = [];
    const source = new U.Menu({items: [
        {caption: 'Axes', value: 'axes', menu: new U.Menu({items: [{caption: 'Horizontal', value: 'bShowHorAxis', checkable: true, stopPropagation: true}]})},
        {caption: 'Legend', value: 'legend', menu: new U.Menu({items: [{caption: 'Left', value: 'LeftLegend', checkable: true, toggleGroup: 'legend', stopPropagation: true}]})}
    ]});
    let props = chartProps(opts.props);
    const api = {asc_registerCallback() {}, asc_getChartPreviews: () => [{asc_getName: () => 1, asc_getImage: () => 'a.png'}, {asc_getName: () => 2, asc_getImage: () => ''}], asc_generateChartPreviews(t, ids) { api.generated = ids; }};
    const host = {
        strings: {txtAddChartElement: 'Add Chart Element', txtQuickLayout: 'Quick Layout', txtChangeColors: 'Change Colors', txtChartStyles: 'Chart Styles',
            txtLayout: 'Layout', txtColorful: 'Colorful', txtMonochromatic: 'Monochromatic'},
        ids: Object.assign({element: 'e', layout: 'l', colors: 'c'}, opts.styles === false ? {} : {styles: 's'}),
        lock: ['x'], transaction: !!opts.transaction,
        add: (id, cmp) => { added[id] = cmp; },
        api: () => api,
        chart: () => props,
        dh: () => ({onChartElement(menu, item) { dhCalls.push(['element', item.value, this.chartProps === props]); }, updateChartElementMenu(menu, p) { dhCalls.push(['update', p === props]); }}),
        sourceMenu: () => source,
        applyStyle: (p, id) => applied.push([p === props, id]),
        done() { host.doneCount = (host.doneCount || 0) + 1; }
    };
    const ribbon = ctx.Common.UI.ChartRibbon.create(host);
    return {ctx, U, ribbon, host, added, applied, dhCalls, api, history, setProps: p => { props = p; }, getProps: () => props};
}

// ---- controls and slots -----------------------------------------------------------------------------
test('the group registers its controls in the slots the tab declares', () => {
    const {added, ribbon} = build();
    assert.deepEqual(Object.keys(added), ['e', 'l', 'c', 's']);
    assert.equal(ribbon.controls.length, 4);
    const none = build({styles: false});
    assert.deepEqual(Object.keys(none.added), ['e', 'l', 'c']);
    assert.equal(none.ribbon.btnStyles, undefined, 'Word keeps its own inline gallery');
});

test('every control has the caption of the host and a ribbon icon', () => {
    const {ribbon} = build();
    assert.equal(ribbon.btnElement.caption, 'Add Chart Element');
    assert.match(ribbon.btnElement.iconCls, /btn-elx-chart-add-element/);
    assert.equal(ribbon.btnLayout.caption, 'Quick Layout');
    assert.match(ribbon.btnLayout.iconCls, /btn-elx-chart-quick-layout/);
    assert.equal(ribbon.btnColors.caption, 'Change Colors');
    assert.match(ribbon.btnColors.iconCls, /btn-elx-chart-colors/);
    assert.equal(ribbon.btnStyles.caption, 'Chart Styles');
    assert.match(ribbon.btnStyles.iconCls, /btn-elx-chart-styles/);
});

// ---- Add Chart Element -----------------------------------------------------------------------------
test('Add Chart Element copies the menu of the document holder on first open and refreshes it each time', () => {
    const {ribbon, dhCalls} = build();
    const menu = ribbon.btnElement.menu;
    menu.trigger('show:before', menu);
    assert.deepEqual(plain(menu.items.map(i => i.options.value)), ['axes', 'legend']);
    const legend = menu.items[1].menu.items[0];
    assert.equal(legend.options.toggleGroup, 'legend');
    assert.equal(legend.options.stopPropagation, true);
    menu.trigger('show:before', menu);
    assert.equal(menu.items.length, 2, 'not copied twice');
    assert.deepEqual(plain(dhCalls.filter(c => c[0] === 'update')), [['update', true], ['update', true]]);
});

test('an element of the menu runs the document holder handler on the chart of the moment', () => {
    const {ribbon, dhCalls, setProps} = build();
    const menu = ribbon.btnElement.menu;
    menu.trigger('show:before', menu);
    setProps(chartProps());
    menu.items[1].menu.items[0].trigger('click');
    assert.deepEqual(plain(dhCalls.filter(c => c[0] === 'element')), [['element', 'LeftLegend', true]]);
});

test('closing the menu gives the focus back to the document', () => {
    const {ribbon, host} = build();
    ribbon.btnElement.menu.trigger('hide:after');
    assert.equal(host.doneCount, 1);
});

// ---- Quick Layout ------------------------------------------------------------------------------------
test('Quick Layout offers eight presets with a thumbnail each', () => {
    const {ctx, ribbon} = build();
    const L = ctx.Common.UI.ChartRibbon.LAYOUTS;
    assert.equal(L.length, 8);
    L.forEach(def => assert.match(ctx.Common.UI.ChartRibbon.thumbnail(def), /^<svg[^>]*>.*<\/svg>$/s));
    const menu = ribbon.btnLayout.menu;
    menu.trigger('show:before', menu);
    assert.ok(menu.inner, 'the gallery is built in the menu');
});

test('a layout sets the title, legend, labels, gridlines and axis titles that differ, in that order', () => {
    const {ctx, getProps} = build();
    const host = {chart: () => getProps(), transaction: false};
    // layout 5: title, legend right, axis titles, gridlines
    ctx.Common.UI.ChartRibbon.applyLayout(host, ctx.Common.UI.ChartRibbon.LAYOUTS[4]);
    const calls = getProps().calls;
    assert.deepEqual(calls.map(c => c[0]), ['legend', 'axisTitles']);
    assert.deepEqual(calls[0].slice(1), [true, 3]);
    assert.deepEqual(calls[1].slice(1), [true, false, true, false, false]);
});

test('a layout that already matches the chart changes nothing', () => {
    const {ctx, getProps} = build();
    const host = {chart: () => getProps()};
    // layout 3: title, legend bottom, gridlines: the chart of the mock already is that
    ctx.Common.UI.ChartRibbon.applyLayout(host, ctx.Common.UI.ChartRibbon.LAYOUTS[2]);
    assert.equal(getProps().calls.length, 0);
});

test('a layout without title, with labels and no gridlines: labels at the end of the bars for a column chart', () => {
    const {ctx, getProps} = build();
    ctx.Common.UI.ChartRibbon.applyLayout({chart: () => getProps()}, ctx.Common.UI.ChartRibbon.LAYOUTS[3]);
    const byName = Object.fromEntries(getProps().calls.map(c => [c[0], c.slice(1)]));
    assert.deepEqual(byName.title, [false, false]);
    assert.deepEqual(byName.labels, [true, 7]);
    assert.equal(byName.grid, undefined, 'the value axis already has major gridlines');
    assert.equal(byName.legend, undefined, 'bottom legend is already there');
});

test('a pie chart gets best-fit labels and no axis steps', () => {
    const {ctx} = build();
    const pie = chartProps({getType: () => 4, getHorAxesProps: () => [], getVertAxesProps: () => [], getLegendPos: () => 3, getTitle: () => 0});
    ctx.Common.UI.ChartRibbon.applyLayout({chart: () => pie}, ctx.Common.UI.ChartRibbon.LAYOUTS[5]);
    const byName = Object.fromEntries(pie.calls.map(c => [c[0], c.slice(1)]));
    assert.deepEqual(byName.labels, [true, 9]);
    assert.deepEqual(byName.title, [true, false]);
    assert.equal(byName.grid, undefined);
    assert.equal(byName.axisTitles, undefined);
});

test('a spreadsheet layout is one undo step: one history point and one transaction', () => {
    const {ctx, history, getProps} = build();
    ctx.Common.UI.ChartRibbon.applyLayout({chart: () => getProps(), transaction: true}, ctx.Common.UI.ChartRibbon.LAYOUTS[4]);
    assert.deepEqual(history.calls, ['point', 'start', 'end']);
});

test('choosing a layout in the gallery applies it and hands the focus back', () => {
    const {ribbon, ctx, getProps, host} = build();
    const menu = ribbon.btnLayout.menu;
    menu.trigger('show:before', menu);
    const picker = menu.inner[0].menu;
    picker.trigger('item:click', picker, {}, {get: k => ({def: ctx.Common.UI.ChartRibbon.LAYOUTS[4]})[k]}, {type: 'click'});
    assert.deepEqual(getProps().calls.map(c => c[0]), ['legend', 'axisTitles']);
    assert.equal(host.doneCount, 1);
});

// ---- Change Colors -----------------------------------------------------------------------------------
test('Change Colors lists Colorful and six monochromatic palettes', () => {
    const {ribbon} = build();
    const items = ribbon.btnColors.menu.items;
    const palettes = items.filter(i => i.options.value !== undefined);
    assert.deepEqual(plain(palettes.map(i => i.options.value)), [-1, 0, 1, 2, 3, 4, 5]);
    assert.equal(palettes[0].options.caption, 'Colorful');
    assert.equal(palettes[3].options.caption, 'Monochromatic 3');
});

test('the palette in use is checked when the menu opens', () => {
    const {ribbon, setProps} = build();
    const menu = ribbon.btnColors.menu;
    setProps(chartProps({getChartColorsAccent: () => 2}));
    menu.trigger('show:before', menu);
    assert.deepEqual(plain(menu.items.filter(i => i.checked).map(i => i.options.value)), [2]);
    setProps(chartProps({getChartColorsAccent: () => -1}));
    menu.trigger('show:before', menu);
    assert.deepEqual(plain(menu.items.filter(i => i.checked).map(i => i.options.value)), [-1]);
});

test('choosing a palette sets it on the chart', () => {
    const {ribbon, getProps, host} = build();
    const menu = ribbon.btnColors.menu;
    menu.trigger('item:click', menu, {value: 4});
    assert.deepEqual(getProps().calls, [['colors', 4]]);
    assert.equal(host.doneCount, 1);
});

// ---- Chart Styles -------------------------------------------------------------------------------------
test('the styles gallery lists the presets of the chart type and asks for their previews', () => {
    const {ribbon, api} = build();
    const menu = ribbon.btnStyles.menu;
    menu.trigger('show:before', menu);
    const picker = menu.inner[0].menu;
    assert.equal(picker.store.models.length, 2);
    assert.deepEqual(plain(api.generated), [1, 2]);
    assert.equal(picker.selected, null, 'style 3 of the chart is not in this mock list');
});

test('a combo chart has no style presets', () => {
    const {ribbon, setProps, api} = build();
    setProps(chartProps({getType: () => 40}));
    const menu = ribbon.btnStyles.menu;
    menu.trigger('show:before', menu);
    assert.equal(menu.inner[0].menu.store.models.length, 0);
    assert.equal(api.generated, undefined);
});

test('choosing a style applies it with the editor call', () => {
    const {ribbon, applied, host} = build();
    const menu = ribbon.btnStyles.menu;
    menu.trigger('show:before', menu);
    const picker = menu.inner[0].menu;
    picker.trigger('item:click', picker, {}, picker.store.models[1], {type: 'click'});
    assert.deepEqual(applied, [[true, 2]]);
    assert.equal(host.doneCount, 1);
});

// ---- wiring in the editors and the strings -----------------------------------------------------------
test('the three editors build the group with the ids their tab declares', () => {
    const word = read('documenteditor/main/app/view/ContextTabs.js');
    assert.match(word, /Common\.UI\.ChartRibbon\.create/);
    for (const id of ['elx-btn-cht-element', 'elx-btn-cht-layout', 'elx-btn-cht-colors']) assert.match(word, new RegExp("big\\('" + id + "'\\)"));
    assert.doesNotMatch(word.slice(word.indexOf('Common.UI.ChartRibbon.create'), word.indexOf('Common.UI.ChartRibbon.create') + 400), /styles:/, 'Word keeps its gallery');
    const excel = read('spreadsheeteditor/main/app/controller/ExcelContextTabs.js');
    assert.match(excel, /Common\.UI\.ChartRibbon\.create/);
    assert.match(excel, /transaction: true/);
    assert.match(excel, /asc_setGraphicObjectProps/);
    for (const id of ['slot-elx-chart-element', 'slot-elx-chart-layout', 'slot-elx-chart-colors', 'slot-elx-chart-styles']) assert.match(excel, new RegExp(id));
    const ppt = read('presentationeditor/main/app/elixir-slide/ContextTabs.js');
    assert.match(ppt, /Common\.UI\.ChartRibbon\.create/);
    assert.match(ppt, /api\.ChartApply|me\.api\.ChartApply/);
    for (const id of ['cd-element', 'cd-layout', 'cd-colors', 'cd-styles']) assert.match(ppt, new RegExp("big\\('" + id + "'\\)"));
    for (const f of [word, excel, ppt]) assert.match(f, /common\/main\/lib\/component\/ChartRibbon/);
});

test('the chart ribbon strings and group labels exist in the nine shipped locales of each editor', () => {
    const spec = {
        documenteditor: ['DE.Views.ContextTabs.', ['txtAddChartElement', 'txtQuickLayout', 'txtChangeColors', 'txtLayout', 'txtColorful', 'txtMonochromatic', 'txtGroupChartlayouts']],
        spreadsheeteditor: ['SSE.Controllers.ExcelContextTabs.', ['txtAddChartElement', 'txtQuickLayout', 'txtChangeColors', 'txtChartStyles', 'txtLayout', 'txtColorful', 'txtMonochromatic']],
        presentationeditor: ['PE.Views.SlideContextTabs.', ['txtAddChartElement', 'txtQuickLayout', 'txtChangeColors', 'txtChartStyles', 'txtLayout', 'txtColorful', 'txtMonochromatic', 'txtGroupChartlayouts', 'txtGroupChartstyles']]
    };
    for (const [ed, [prefix, keys]] of Object.entries(spec)) {
        const en = JSON.parse(read(`${ed}/main/locale/en.json`));
        for (const l of LOCALES) {
            const d = JSON.parse(read(`${ed}/main/locale/${l}.json`));
            for (const k of keys) {
                assert.ok(d[prefix + k], `${ed} ${l} ${prefix}${k}`);
                if (l !== 'en' && !['txtLayout'].includes(k)) assert.notEqual(d[prefix + k], en[prefix + k] && k !== 'txtMonochromatic' ? en[prefix + k] : null, `${ed} ${l} ${k} is translated`);
            }
        }
    }
    for (const l of LOCALES) {
        const d = JSON.parse(read(`spreadsheeteditor/main/locale/${l}.json`));
        assert.ok(d['SSE.Views.RibbonTabs.txtGroupChartlayouts'] && d['SSE.Views.RibbonTabs.txtGroupChartstyles'], 'Excel group labels ' + l);
    }
});

test('the presentation editor chart element menu labels Gridlines', () => {
    const view = read('presentationeditor/main/app/view/DocumentHolderExt.js');
    assert.doesNotMatch(view, /me\.textGridLines\b/);
});

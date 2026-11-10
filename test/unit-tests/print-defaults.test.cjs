// SPDX-License-Identifier: AGPL-3.0-only
'use strict';
// Spreadsheet print panel: smart defaults decision logic, notices and the "not saved unless the user saves" rule.
const test = require('node:test'), assert = require('node:assert/strict'), fs = require('node:fs'), path = require('node:path'), vm = require('node:vm');
const root = path.resolve(__dirname, '../..');
const PD = require(path.join(root, 'apps/common/main/lib/elixir-local/PrintDefaults.js'));
const A4 = {pageWidth: 210, pageHeight: 297, margins: {left: 17.8, right: 17.8, top: 19.1, bottom: 19.1}};
const plain = {scale: 100, fitToWidth: 0, fitToHeight: 0, orientation: 0, horizontalCentered: false};
const decide = o => PD.decide(Object.assign({stored: plain, hasPrintArea: false}, A4, o));

test('a range that fits the page at 100% is left alone', () => {
    const d = decide({rangeWidth: 150, rangeHeight: 100});
    assert.deepEqual([d.fitToWidth, d.orientation, d.horizontalCentered], [false, undefined, false]);
});
test('a range too big for 100% but readable on one page gets 1 x 1 (portrait when that is best), centred', () => {
    let d = decide({rangeWidth: 300, rangeHeight: 500});       // portrait scale 51.8%
    assert.equal(d.fitToWidth, true); assert.equal(d.fitToHeight, true); assert.equal(d.orientation, undefined); assert.equal(d.horizontalCentered, true);
    assert.deepEqual(d.reasons, ['one-page', 'center']);
    d = decide({rangeWidth: 200, rangeHeight: 270});           // slightly too long: 87%
    assert.equal(d.fitToHeight, true); assert.equal(d.orientation, undefined); assert.ok(d.scale > 0.85 && d.scale < 0.9);
});
test('the better orientation is chosen for one page: sideways when it gives the larger scale', () => {
    let d = decide({rangeWidth: 240, rangeHeight: 100});       // portrait 73%, landscape 100%+
    assert.equal(d.orientation, PD.LANDSCAPE); assert.equal(d.fitToHeight, true);
    d = decide({rangeWidth: 400, rangeHeight: 100});           // landscape 65%
    assert.equal(d.orientation, PD.LANDSCAPE); assert.equal(d.fitToHeight, true); assert.equal(d.horizontalCentered, true);
    d = decide({rangeWidth: 150, rangeHeight: 300});           // portrait 86%, landscape worse
    assert.equal(d.orientation, undefined); assert.equal(d.fitToHeight, true);
});
test('a sheet that already fits one page at 100% is left alone', () => {
    const d = decide({rangeWidth: 150, rangeHeight: 100});
    assert.equal(d.fitToWidth, false); assert.equal(d.fitToHeight, false);
});
test('a huge sheet below 45% on one page keeps fit-to-width (1 x 0) across several pages', () => {
    let d = decide({rangeWidth: 800, rangeHeight: 2000});
    assert.equal(d.fitToWidth, true); assert.equal(d.fitToHeight, false); assert.equal(d.orientation, undefined); assert.deepEqual(d.reasons, ['fit-width', 'center']);
    d = decide({rangeWidth: 1500, rangeHeight: 300});          // wide and short: sideways, still several pages wide-fitted
    assert.equal(d.fitToHeight, false); assert.equal(d.orientation, PD.LANDSCAPE);
    assert.ok(PD.MIN_ONE_PAGE_SCALE === 45);
});
test('explicit settings are respected: scale, fit-to, landscape, centring', () => {
    assert.equal(decide({stored: Object.assign({}, plain, {scale: 80}), rangeWidth: 400, rangeHeight: 100}).fitToWidth, false);
    assert.equal(decide({stored: Object.assign({}, plain, {fitToHeight: 1}), rangeWidth: 400, rangeHeight: 100}).fitToWidth, false);
    let d = decide({stored: Object.assign({}, plain, {orientation: 1}), rangeWidth: 400, rangeHeight: 100});
    assert.equal(d.orientation, undefined); assert.equal(d.fitToHeight, true);
    d = decide({stored: Object.assign({}, plain, {orientation: 1}), rangeWidth: 800, rangeHeight: 600});   // landscape scale 32%: width only
    assert.equal(d.fitToHeight, false); assert.equal(d.fitToWidth, true);
    d = decide({stored: Object.assign({}, plain, {horizontalCentered: true}), rangeWidth: 300, rangeHeight: 500});
    assert.equal(d.horizontalCentered, false);
});
test('a user print area, unknown sizes and empty ranges change nothing', () => {
    for (const extra of [{hasPrintArea: true, rangeWidth: 400, rangeHeight: 100}, {rangeWidth: null, rangeHeight: 100}, {rangeWidth: 400, rangeHeight: undefined}, {rangeWidth: 0, rangeHeight: 0}, {pageWidth: null, rangeWidth: 400, rangeHeight: 100}]) {
        const d = decide(extra); assert.equal(d.fitToWidth, false); assert.equal(d.orientation, undefined);
    }
});
test('margins reported by the engine are what counts; oversized ones cannot create a negative area', () => {
    assert.equal(decide({margins: {left: 5, right: 5}, rangeWidth: 195, rangeHeight: 100}).fitToWidth, false);
    assert.equal(decide({margins: {left: 5, right: 5}, rangeWidth: 205, rangeHeight: 100}).fitToHeight, true);
    assert.equal(decide({margins: {left: 150, right: 150}, rangeWidth: 205, rangeHeight: 300}).fitToWidth, false);
});
test('a stored scale below 25% without fit-to gets one page when readable, else the page width', () => {
    const tiny = Object.assign({}, plain, {scale: 10});
    let d = decide({stored: tiny, rangeWidth: 100, rangeHeight: 100});
    assert.equal(d.fitToWidth, true); assert.equal(d.fitToHeight, true); assert.equal(d.tinyScale, true); assert.equal(d.horizontalCentered, true); assert.deepEqual(d.reasons, ['one-page', 'center']);
    d = decide({stored: tiny, rangeWidth: 800, rangeHeight: 2000});
    assert.equal(d.fitToWidth, true); assert.equal(d.fitToHeight, false); assert.equal(d.tinyScale, true);
    d = decide({stored: tiny, rangeWidth: null, rangeHeight: null, pageWidth: null});   // no size information from the engine
    assert.equal(d.fitToWidth, true); assert.equal(d.fitToHeight, false); assert.equal(d.tinyScale, true); assert.deepEqual(d.reasons, ['tiny-scale', 'center']);
    d = decide({stored: tiny, hasPrintArea: true, rangeWidth: 400, rangeHeight: 100});  // a print area is a user range: width only
    assert.equal(d.fitToWidth, true); assert.equal(d.fitToHeight, false); assert.equal(d.orientation, undefined);
});
test('tiny scale: landscape and centring rules are unchanged', () => {
    const tiny = Object.assign({}, plain, {scale: 10});
    let d = decide({stored: tiny, rangeWidth: 400, rangeHeight: 100});
    assert.equal(d.orientation, PD.LANDSCAPE); assert.equal(d.fitToWidth, true);
    d = decide({stored: Object.assign({}, tiny, {horizontalCentered: true}), rangeWidth: 400, rangeHeight: 100}); assert.equal(d.horizontalCentered, false);
    d = decide({stored: Object.assign({}, tiny, {orientation: 1}), rangeWidth: 400, rangeHeight: 100}); assert.equal(d.orientation, undefined);
});
test('a stored 50% scale, 25% and a fit-to already set are left alone', () => {
    for (const stored of [{scale: 50}, {scale: 25}, {scale: 10, fitToWidth: 1}, {scale: 10, fitToHeight: 1}]) {
        const d = decide({stored: Object.assign({}, plain, stored), rangeWidth: 100, rangeHeight: 100});
        assert.equal(d.fitToWidth, false, JSON.stringify(stored)); assert.equal(d.tinyScale, false);
    }
    assert.equal(PD.hasTinyScale({scale: 24, fitToWidth: 0, fitToHeight: 0}), true);
    assert.equal(PD.hasTinyScale({scale: 24, fitToWidth: 1}), false);
});
test('warning messages are filled in and engine results normalised', () => {
    const t = {scale: 'Sheet {name} is set to {scale}%', margins: 'Sheet {name} margins'};
    assert.equal(PD.formatWarning({kind: 'scale', name: 'Data', scale: 17.6}, t), 'Sheet Data is set to 18%');
    assert.equal(PD.formatWarning({kind: 'margins', name: 'Data'}, t), 'Sheet Data margins');
    assert.equal(PD.formatWarning({kind: 'fitted', name: 'Data', scale: 10}, Object.assign({fitted: 'Sheet {name} was stored at {scale}%'}, t)), 'Sheet Data was stored at 10%');
    assert.deepEqual(PD.normalizeWarnings([{sheetIndex: 1, name: 'A', scale: 10}, null, {kind: 'margins', sheetIndex: 2, name: 'B'}]).map(w => w.kind), ['scale', 'margins']);
    assert.deepEqual(PD.normalizeWarnings(undefined), []);
});

// ---- controller -------------------------------------------------------------------------------------------------
function jq() {   // minimal chainable jQuery stand-in
    const make = () => { const el = {kids: [], texts: [], hiddenClass: null, handlers: {}, prepend() { return el; }, append(...k) { el.kids.push(...k); return el; }, empty() { el.kids = []; return el; },
        toggleClass(n, on) { el.hiddenClass = on; return el; }, text(t) { el.texts.push(t); el.label = t; return el; }, on(n, f) { el.handlers[n] = f; return el; }}; return el; };
    const roots = [];
    const $ = () => { const e = make(); roots.push(e); return e; };
    $.roots = roots; return $;
}
function load() {
    const $ = jq(), saved = [];
    const context = {console, SSE: {Controllers: {}, getController: () => ({getSelectTabs: () => context.__tabs || [0]})}, Asc: {referenceType: {A: 0}, c_oAscPrintType: {ActiveSheets: 0, EntireWorkbook: 1, Selection: 2}, asc_CAdjustPrint: function () { this.setType = t => this.type = t; this.asc_setPrintType = t => this.t = t; this.asc_setIgnorePrintArea = v => this.i = v; }},
        _: {extend: Object.assign, bind: (f, t) => f.bind(t), findWhere: (list, q) => list.find(x => Object.keys(q).every(k => x[k] === q[k])), each: (l, f) => l.forEach(f)},
        Backbone: {Controller: {extend: spec => spec}}, $, Common: {NotificationCenter: {trigger() {}}}, define(_deps, factory) { context.__ctrl = factory({}, PD); }};
    vm.runInNewContext(fs.readFileSync(path.join(root, 'apps/spreadsheeteditor/main/app/controller/Print.js'), 'utf8'), context);
    const c = context.SSE.Controllers.Print;
    return {c, saved, $};
}
const pageOptions = (o = {}) => { const setup = Object.assign({scale: 100, fitToWidth: 0, fitToHeight: 0, orientation: 0, width: 210, height: 297}, o.setup), st = {hc: !!o.hc};
    return {setup, st, asc_getPageSetup: () => ({asc_getScale: () => setup.scale, asc_getFitToWidth: () => setup.fitToWidth, asc_getFitToHeight: () => setup.fitToHeight, asc_getOrientation: () => setup.orientation,
        asc_getWidth: () => setup.width, asc_getHeight: () => setup.height, asc_setFitToWidth: v => setup.fitToWidth = v, asc_setFitToHeight: v => setup.fitToHeight = v, asc_setScale: v => setup.scale = v, asc_setOrientation: v => setup.orientation = v}),
        asc_getPageMargins: () => ({asc_getLeft: () => 17.8, asc_getRight: () => 17.8, asc_getTop: () => 19.1, asc_getBottom: () => 19.1}), asc_getHorizontalCentered: () => st.hc, asc_setHorizontalCentered: v => st.hc = v}; };
function panelFixture(info, local = true) {
    const e = load(), c = e.c, stored = [], saved = [];
    c.mode = {isElixirLocal: local}; c.printSettings = {getIgnorePrintArea: () => false, pnlSettings: {prepend() {}}, cmbSheet: {getValue: () => 0}};
    c._changedProps = []; c._storedProps = []; c._smartPending = {};
    c.api = {asc_GetPrintRangeInfo: () => info, asc_getPageOptions: () => pageOptions(), asc_savePagePrintOptions: map => saved.push(map), wbModel: {getDefinesNames: () => null, getWorksheet: () => ({getId: () => 1})}};
    return Object.assign(e, {c, saved});
}
test('smart defaults shape the panel state for a wide sheet and the preview/PDF share that same options object', () => {
    const f = panelFixture([{sheetIndex: 0, widthMm: 500, heightMm: 120, usedMargins: {left: 17.8, right: 17.8}}]), c = f.c;
    c._changedProps[0] = pageOptions(); c.applySmartDefaults(0);
    assert.equal(c._changedProps[0].setup.fitToWidth, 1); assert.equal(c._changedProps[0].setup.fitToHeight, 1, 'one page per sheet');
    assert.equal(c._changedProps[0].setup.orientation, 1); assert.equal(c._changedProps[0].st.hc, true);
    assert.equal(c._smartPending[0].orientation, true);
});
test('a stored 10% scale is fitted in the panel state and shown as information, not saved into the document', () => {
    const f = panelFixture([{sheetIndex: 0, widthMm: 120, heightMm: 100}]), c = f.c, seen = [];
    c.txtFittedInfo = 'Sheet {name} was stored at {scale}%'; c.txtScaleWarning = 'S'; c.txtMarginsWarning = 'M'; c.txtFitToWidth = 'Fit';
    c.api.asc_getWorksheetName = () => 'Data'; c.api.asc_GetPrintSetupWarnings = () => [];
    c.api.asc_getPageOptions = () => pageOptions({setup: {scale: 10}});   // the document's own copy
    c._changedProps[0] = pageOptions({setup: {scale: 10}}); c.applySmartDefaults(0);
    const setup = c._changedProps[0].setup;
    assert.equal(setup.fitToWidth, 1); assert.equal(setup.fitToHeight, 1); assert.equal(setup.scale, 100);
    assert.equal(c._smartPending[0].tinyScale, 10);
    c.updateScaleWarning({});
    assert.equal(c._scaleWarning.kids.length, 1); assert.equal(c._scaleWarning.kids[0].kids[0].label, 'Sheet Data was stored at 10%');
    assert.equal(c._scaleWarning.kids[0].kids.length, 1, 'information has no action');
    c.savePageOptions(null); assert.equal(f.saved[0][0].setup.scale, 10, 'the stored value is untouched in the document');
});
test('a stored 50% scale is left alone by the panel', () => {
    const f = panelFixture([{sheetIndex: 0, widthMm: 120, heightMm: 100}]); f.c._changedProps[0] = pageOptions({setup: {scale: 50}}); f.c.applySmartDefaults(0);
    assert.equal(f.c._changedProps[0].setup.fitToWidth, 0); assert.equal(f.c._smartPending[0], undefined);
});
test('smart defaults are off outside the local desktop and without engine size information', () => {
    let f = panelFixture([{sheetIndex: 0, widthMm: 500, heightMm: 120}], false); f.c._changedProps[0] = pageOptions(); f.c.applySmartDefaults(0); assert.equal(f.c._changedProps[0].setup.fitToWidth, 0);
    f = panelFixture([{sheetIndex: 0}]); f.c._changedProps[0] = pageOptions(); f.c.applySmartDefaults(0); assert.equal(f.c._changedProps[0].setup.fitToWidth, 0);
    f = panelFixture([{sheetIndex: 0, widthMm: 500, heightMm: 120, hasPrintArea: true}]); f.c._changedProps[0] = pageOptions(); f.c.applySmartDefaults(0); assert.equal(f.c._changedProps[0].setup.fitToWidth, 0);
});
test('print and Save as PDF do not write smart defaults into the document; the panel Save and user edits do', () => {
    const f = panelFixture([{sheetIndex: 0, widthMm: 500, heightMm: 120}]), c = f.c;
    c._changedProps[0] = pageOptions(); c._changedProps[1] = pageOptions({setup: {scale: 80}});
    c.applySmartDefaults(0);
    c.savePageOptions(null);
    assert.equal(f.saved[0][0].setup.fitToWidth, 0, 'document keeps its stored setup for the smart sheet');
    assert.equal(f.saved[0][1].setup.scale, 80, 'other sheets are saved as before');
    c.savePageOptions(null, true); assert.equal(f.saved[1][0].setup.fitToWidth, 1, 'the panel Save keeps what the panel shows');
    delete c._smartPending[0]; c.savePageOptions(null); assert.equal(f.saved[2][0].setup.fitToWidth, 1, 'a sheet edited by the user is saved');
});
test('notices: setup warnings are preferred, scale-only API is the fallback, Fit to width acts on the sheet', () => {
    const f = panelFixture([]), c = f.c; let called = [];
    c.txtScaleWarning = 'S {name} {scale}'; c.txtMarginsWarning = 'M {name}'; c.txtFitToWidth = 'Fit';
    c.api.asc_GetPrintSetupWarnings = () => { called.push('setup'); return [{kind: 'scale', sheetIndex: 0, name: 'Data', scale: 12}, {kind: 'margins', sheetIndex: 0, name: 'Data', margins: {}, usedMargins: {}}]; };
    c.api.asc_GetPrintScaleWarnings = () => { called.push('scale'); return []; };
    c.updateScaleWarning({}); assert.deepEqual(called, ['setup']);
    const rows = c._scaleWarning.kids; assert.equal(rows.length, 2);
    assert.equal(rows[0].kids[0].label, 'S Data 12'); assert.equal(rows[1].kids[0].label, 'M Data'); assert.equal(rows[1].kids.length, 1, 'margins notice has no action');
    const link = rows[0].kids.find(k => k.handlers && k.handlers.click); assert.equal(link.label, 'Fit');
    let updated = 0, filled = 0; c.updatePreview = () => updated++; c.fillPageOptions = () => filled++;
    c._changedProps[0] = pageOptions({setup: {scale: 12}}); c._smartPending[0] = {};
    link.handlers.click({preventDefault() {}});
    assert.equal(c._changedProps[0].setup.fitToWidth, 1); assert.equal(c._changedProps[0].setup.fitToHeight, 0); assert.equal(c._changedProps[0].setup.scale, 100);
    assert.equal(updated, 1); assert.equal(filled, 1); assert.equal(c._smartPending[0], undefined);
    delete c.api.asc_GetPrintSetupWarnings; called = []; c.updateScaleWarning({}); assert.deepEqual(called, ['scale']);
});

// ---- print range: sheets, labels, empty selection ------------------------------------------------------------------
function rangeFixture(o = {}) {
    const f = panelFixture([]), c = f.c, state = {disabled: [], preview: null, empty: null, text: null};
    c.api.asc_getWorksheetsCount = () => 4; c.api.asc_isWorksheetHidden = i => (o.hidden || []).includes(i); c.api.asc_getWorksheetName = i => 'S' + i;
    c.api.asc_getActiveRangeStr = () => o.range === undefined ? 'A1:C3' : o.range;
    c.api.asc_getCellInfo = () => ({asc_getText: () => o.text || '', asc_getFormula: () => ''});
    c._checkedSheets = o.checked || {};
    const btn = () => ({setDisabled: d => state.disabled.push(d)});
    Object.assign(c.printSettings, {getRange: () => o.type === undefined ? 1 : o.type, isChooseSheets: () => !!o.choose,
        txtEmptyTable: 'table empty', txtEmptySelection: 'The selection is empty', txtNoSheetChosen: 'Choose at least one sheet',
        $previewBox: {toggleClass: (n, on) => { state.preview = on; }},
        $previewEmpty: {toggleClass: (n, on) => { state.empty = !on; state.hiddenEmpty = on; return {children: () => ({text: t => { state.text = t; }})}; }},
        btnsPrint: [btn()], btnsPrintPDF: [btn()]});
    return Object.assign(f, {state});
}
test('Choose sheets prints exactly the ticked visible sheets as active sheets', () => {
    let f = rangeFixture({choose: true, hidden: [2], checked: {1: false}});
    assert.deepEqual(Array.from(f.c.getActiveSheetsArray(f.c.printSettings)), [0, 3]);
    f = rangeFixture({type: 0});                      // Active sheet: the selected tabs, as before
    assert.deepEqual(Array.from(f.c.getActiveSheetsArray(f.c.printSettings)), [0]);
    f = rangeFixture({type: 1});                      // Entire workbook: no array, the engine takes every sheet
    assert.equal(f.c.getActiveSheetsArray(f.c.printSettings), null);
});
test('an empty or single-empty-cell selection shows the notice and disables Print and Save as PDF', () => {
    let f = rangeFixture({type: 2, range: 'B2', text: ''}); assert.equal(f.c.showPreviewState(3), 0);
    assert.equal(f.state.text, 'The selection is empty'); assert.equal(f.state.preview, true); assert.equal(f.state.hiddenEmpty, false); assert.deepEqual(f.state.disabled, [true, true]);
    f = rangeFixture({type: 2, range: 'A1:C3', text: ''}); assert.equal(f.c.showPreviewState(2), 2); assert.deepEqual(f.state.disabled, [false, false]);
    f = rangeFixture({type: 2, range: 'B2', text: 'x'}); assert.equal(f.c.showPreviewState(1), 1);
    f = rangeFixture({type: 2, range: 'A1:C3'}); assert.equal(f.c.showPreviewState(0), 0); assert.equal(f.state.text, 'The selection is empty');
});
test('no ticked sheet: notice and disabled buttons; an empty table keeps its own message', () => {
    let f = rangeFixture({choose: true, checked: {0: false, 1: false, 2: false, 3: false}}); assert.equal(f.c.showPreviewState(4), 0);
    assert.equal(f.state.text, 'Choose at least one sheet'); assert.deepEqual(f.state.disabled, [true, true]);
    f = rangeFixture({type: 1}); assert.equal(f.c.showPreviewState(0), 0); assert.equal(f.state.text, 'table empty');
    f = rangeFixture({type: 1}); assert.equal(f.c.showPreviewState(5), 5); assert.deepEqual(f.state.disabled, [false, false]);
});
test('the range entries are named for the desktop: entire workbook first, grouped tabs as Selected sheets (N)', () => {
    const src = fs.readFileSync(path.join(root, 'apps/spreadsheeteditor/main/app/view/FileMenuPanels.js'), 'utf8'), ctl = fs.readFileSync(path.join(root, 'apps/spreadsheeteditor/main/app/controller/Print.js'), 'utf8');
    assert.match(src, /txtSelectedSheets\.replace\('\{n\}', selectedTabs\)/); assert.match(src, /txtActiveSheets: 'Active sheet'/); assert.match(src, /txtSelection: 'Selected cells'/);
    assert.match(src, /txtAllSheets: 'Entire workbook'/); assert.match(src, /txtChooseSheets: 'Choose sheets…'/);
    assert.match(ctl, /updateRangeLabels\(SSE\.getController\('Statusbar'\)\.getSelectTabs\(\)\.length\)/);
    const en = fs.readFileSync(path.join(root, 'apps/spreadsheeteditor/main/locale/en.json'), 'utf8'), fr = fs.readFileSync(path.join(root, 'apps/spreadsheeteditor/main/locale/fr.json'), 'utf8');
    for (const [k, e, r] of [['txtAllSheets', 'Entire workbook', 'Classeur entier'], ['txtActiveSheets', 'Active sheet', 'Feuille active'], ['txtChooseSheets', 'Choose sheets…', 'Choisir les onglets…'],
        ['txtSelection', 'Selected cells', 'Cellules sélectionnées'], ['txtSelectedSheets', 'Selected sheets ({n})', 'Feuilles sélectionnées ({n})'], ['txtEmptySelection', 'The selection is empty', 'La sélection est vide']]) {
        assert.ok(en.includes(`"SSE.Views.PrintWithPreview.${k}": "${e}"`), 'en ' + k); assert.ok(fr.includes(`"SSE.Views.PrintWithPreview.${k}": "${r}"`), 'fr ' + k);
    }
});

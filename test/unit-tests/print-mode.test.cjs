// SPDX-License-Identifier: AGPL-3.0-only
'use strict';
// Print panel: one tool, mode by entry point; default print range; button labels and layout.
const test = require('node:test'), assert = require('node:assert/strict'), fs = require('node:fs'), path = require('node:path');
const root = path.resolve(__dirname, '../..'), read = f => fs.readFileSync(path.join(root, f), 'utf8');
const PM = require(path.join(root, 'apps/common/main/lib/elixir-local/PrintMode.js'));

function settings(initial) { const v = {}; if (initial) v[PM.KEY] = initial; return {get: k => v[k], set: (k, x) => { v[k] = x; }, v}; }
function button(cls) { const el = {classes: new Set(cls), toggleClass(n, on) { on ? el.classes.add(n) : el.classes.delete(n); }}; return {cmpEl: el}; }
function view(arrays) {
    const title = {texts: [], length: 1, text(t) { title.texts.push(t); }}, print = button(['primary']), pdf = button([]);
    return arrays ? {$el: {find: () => title}, btnsPrint: [print, button(['primary'])], btnsPrintPDF: [pdf, button([])], title, print, pdf}
        : {$el: {find: () => title}, btnPrint: print, btnPrintPdf: pdf, title, print, pdf};
}
const strings = {print: 'Print', exportPdf: 'Export as PDF'};

test('the requested mode is consumed once: a later Ctrl+P is a plain Print', () => {
    const s = settings('pdf');
    assert.equal(PM.consume(s), 'pdf'); assert.equal(PM.consume(s), 'print'); assert.equal(PM.consume(settings()), 'print');
    assert.equal(PM.consume(null), 'print');
});
test('Export as PDF mode: title and primary button; Print mode: the opposite (single buttons)', () => {
    let v = view(false); PM.apply(v, 'pdf', strings);
    assert.equal(v.title.texts.at(-1), 'Export as PDF'); assert.equal(v.print.cmpEl.classes.has('primary'), false); assert.equal(v.pdf.cmpEl.classes.has('primary'), true);
    PM.apply(v, 'print', strings);
    assert.equal(v.title.texts.at(-1), 'Print'); assert.equal(v.print.cmpEl.classes.has('primary'), true); assert.equal(v.pdf.cmpEl.classes.has('primary'), false);
});
test('the spreadsheet panel (two button rows) follows the same mode', () => {
    const v = view(true); PM.apply(v, 'pdf', strings);
    assert.ok(v.btnsPrint.every(b => !b.cmpEl.classes.has('primary'))); assert.ok(v.btnsPrintPDF.every(b => b.cmpEl.classes.has('primary')));
});
test('all three Print controllers read the mode when the panel is shown, only in the local desktop app', () => {
    for (const ed of ['documenteditor', 'presentationeditor', 'spreadsheeteditor']) {
        const src = read(`apps/${ed}/main/app/controller/Print.js`);
        assert.match(src, /elixir-local\/PrintMode/);
        assert.match(src, /if \(this\.mode && this\.mode\.isElixirLocal\)\s+PrintMode\.apply\(this\.printSettings, PrintMode\.consume\(Common\.Utils\.InternalSettings\)/);
    }
});
test('the spreadsheet panel defaults to the entire workbook in the desktop app, before the panel is filled', () => {
    const src = read('apps/spreadsheeteditor/main/app/controller/Print.js');
    const a = src.indexOf('this.adjPrintParams.asc_setPrintType(Asc.c_oAscPrintType.EntireWorkbook)'), b = src.indexOf('this.fillPrintOptions(this.adjPrintParams, false);');
    assert.ok(a > 0 && b > a, 'default set before fillPrintOptions'); assert.match(src.slice(a - 120, a), /isElixirLocal/);
});
test('labels: Export as PDF title, Save as PDF, Save settings in English and French', () => {
    const want = {documenteditor: ['DE', 'txtPrintPdf'], presentationeditor: ['PE', 'txtPrintPdf'], spreadsheeteditor: ['SSE', 'txtPrintToPDF']};
    for (const [ed, [ns, pdfKey]] of Object.entries(want)) {
        const en = read(`apps/${ed}/main/locale/en.json`), fr = read(`apps/${ed}/main/locale/fr.json`), view = read(`apps/${ed}/main/app/view/FileMenuPanels.js`);
        assert.match(en, new RegExp(`"${ns}.Views.PrintWithPreview.txtExportPdf": "Export as PDF"`)); assert.match(fr, new RegExp(`"${ns}.Views.PrintWithPreview.txtExportPdf": "Exporter en PDF"`));
        assert.match(en, new RegExp(`"${ns}.Views.PrintWithPreview.${pdfKey}": "Save as PDF"`)); assert.match(view, /txtExportPdf: 'Export as PDF'/);
    }
    assert.match(read('apps/spreadsheeteditor/main/locale/en.json'), /"SSE.Views.PrintWithPreview.txtSave": "Save settings"/);
    assert.match(read('apps/spreadsheeteditor/main/locale/fr.json'), /"SSE.Views.PrintWithPreview.txtSave": "Enregistrer les réglages"/);
    assert.match(read('apps/spreadsheeteditor/main/app/view/FileMenuPanels.js'), /txtSave: 'Save settings'/);
});
test('the button row wraps at any panel width (no fixed-width table), in the three editors', () => {
    const view = read('apps/spreadsheeteditor/main/app/view/FileMenuPanels.js');
    assert.match(view, /<div class="print-btn-row">/); assert.doesNotMatch(view.slice(view.indexOf('print-btn-row') - 300, view.indexOf('print-btn-row') + 900), /<table>/);
    const sse = read('apps/spreadsheeteditor/main/resources/less/leftmenu.less').replace(/\r\n/g, '\n');
    assert.match(sse, /\.print-btn-row \{\s+display: flex;\s+flex-wrap: wrap;/); assert.doesNotMatch(sse, /\.footer-with-pdf \{\s+table \{\s+width: 247px/);
    for (const f of ['apps/documenteditor/main/resources/less/filemenu.less', 'apps/presentationeditor/main/resources/less/leftmenu.less'])
        assert.match(read(f).replace(/\r\n/g, '\n'), /\.footer \{\s+margin-top: 24px;\s+(?:\n\s*)?display: flex;\s+flex-wrap: wrap;/);
});

// SPDX-License-Identifier: AGPL-3.0-only
'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const root = path.resolve(__dirname, '../..');
const sse = path.join(root, 'apps/spreadsheeteditor/main');
const underscore = require(path.join(root, 'vendor/underscore/underscore-node.cjs'));

function evalFile(file, extra = {}) {
    const context = {
        console, _: underscore, SSE: {Controllers: {}, Views: {}}, Asc: {}, window: {},
        Backbone: {Controller: {extend: spec => spec}},
        Common: {util: {}, Utils: {}, NotificationCenter: {trigger() {}}},
        define(_deps, factory) { factory(); },
        ...extra
    };
    vm.runInNewContext(fs.readFileSync(file, 'utf8'), context, {filename: path.basename(file)});
    return context;
}

// ---- Workbook Statistics --------------------------------------------------------------------------
function cell(text, formula) { return {isNullText: () => !text && !formula, formulaParsed: formula ? {} : null}; }
function sheet(opts) {
    const cells = opts.cells || [];
    return {
        getHidden: () => !!opts.hidden, getRowsCount: () => opts.rows === undefined ? 10 : opts.rows, getColsCount: () => 5,
        getRange3: () => ({_foreachNoEmpty: fn => cells.forEach(c => fn(c))}),
        TableParts: new Array(opts.tables || 0), pivotTables: new Array(opts.pivots || 0), aComments: new Array(opts.comments || 0),
        Drawings: (opts.drawings || []).map(t => ({graphicObject: {getObjectType: () => t}}))
    };
}

test('Workbook Statistics counts the whole workbook', () => {
    const ctx = evalFile(path.join(sse, 'app/controller/ExcelReview.js'), {AscDFH: {historyitem_type_ChartSpace: 7, historyitem_type_ImageShape: 9}});
    const er = ctx.SSE.Controllers.ExcelReview;
    const wb = {aWorksheets: [
        sheet({cells: [cell('a'), cell('', true), cell('b'), cell('')], tables: 2, pivots: 1, comments: 3, drawings: [7, 7, 9, 1]}),
        sheet({hidden: true, cells: [cell('x', true)], comments: 1, drawings: [9]}),
        sheet({rows: 0})
    ]};
    assert.deepEqual(JSON.parse(JSON.stringify(er.collectStats(wb))),
        {sheets: 3, hidden: 1, cells: 4, formulas: 2, tables: 2, pivots: 1, charts: 2, pictures: 2, shapes: 1, comments: 4});
});

test('Workbook Statistics of an empty workbook is all zero', () => {
    const ctx = evalFile(path.join(sse, 'app/controller/ExcelReview.js'), {AscDFH: {}});
    const r = ctx.SSE.Controllers.ExcelReview.collectStats({aWorksheets: [sheet({rows: 0})]});
    assert.equal(r.sheets, 1);
    assert.equal(r.cells + r.formulas + r.charts + r.pictures + r.shapes, 0);
});

// ---- Previous / Next comment ----------------------------------------------------------------------
function commentFixture(items) {
    const ctx = evalFile(path.join(sse, 'app/controller/ExcelReview.js'));
    const er = Object.assign({}, ctx.SSE.Controllers.ExcelReview);
    const calls = [];
    const models = items.map(i => ({get: k => ({uid: i.uid, hide: i.hide, removed: i.removed})[k]}));
    const collection = {each: fn => models.forEach(fn)};
    const pos = {};
    items.forEach(i => { pos[i.uid] = i; });
    const api = {
        asc_findComment: uid => pos[uid] && pos[uid].row !== undefined ? {asc_getRow: () => pos[uid].row, asc_getCol: () => pos[uid].col} : null,
        asc_findCell: ref => calls.push(['cell', ref]),
        asc_selectComment: id => calls.push(['select', id]), asc_showComment: id => calls.push(['show', id])
    };
    er.getApi = () => api;
    er.getApplication = () => ({getController: () => ({collection})});
    er._lastComment = null;
    return {er, calls, api, collection};
}

test('comments of the worksheet are ordered row by row, hidden ones and other sheets are skipped', () => {
    const f = commentFixture([{uid: 'c', row: 9, col: 0}, {uid: 'a', row: 2, col: 5}, {uid: 'b', row: 2, col: 1},
        {uid: 'h', row: 0, col: 0, hide: true}, {uid: 'o'}, {uid: 'r', row: 1, col: 1, removed: true}]);
    assert.deepEqual(f.er.commentList(f.api, f.collection).map(c => c.id), ['b', 'a', 'c']);
});

test('Next and Previous walk the comments and wrap around', () => {
    const f = commentFixture([{uid: 'c', row: 9, col: 0}, {uid: 'a', row: 2, col: 5}, {uid: 'b', row: 2, col: 1}]);
    const walk = dir => { f.calls.length = 0; f.er.gotoComment(dir); return f.calls.find(c => c[0] === 'select')[1]; };
    assert.deepEqual([1, 1, 1, 1].map(walk), ['b', 'a', 'c', 'b']);
    assert.deepEqual([-1, -1, -1].map(walk), ['c', 'a', 'b']);
    assert.deepEqual(f.calls.map(c => c[0]), ['cell', 'select', 'show']);
});

test('Next and Previous select the commented cell first (like Excel), then show its note', () => {
    const f = commentFixture([{uid: 'c', row: 9, col: 0}, {uid: 'a', row: 2, col: 27}, {uid: 'b', row: 2, col: 1}]);
    const cells = dir => { f.calls.length = 0; f.er.gotoComment(dir); return f.calls.filter(c => c[0] === 'cell').map(c => c[1]); };
    assert.deepEqual([1, 1, 1, 1].map(cells), [['B3'], ['AB3'], ['A10'], ['B3']]);
    assert.deepEqual(cells(-1), [['A10']].flat());
    f.calls.length = 0;
    f.er.gotoComment(1);
    assert.deepEqual(f.calls.map(c => c[0]), ['cell', 'select', 'show']);
});

test('cell references are letters + row, columns past Z carry over', () => {
    const f = commentFixture([]);
    assert.deepEqual([[0, 0], [0, 25], [0, 26], [0, 51], [0, 52], [4, 701], [0, 702], [1048575, 16383]].map(p => f.er.cellRef(p[0], p[1])),
        ['A1', 'Z1', 'AA1', 'AZ1', 'BA1', 'ZZ5', 'AAA1', 'XFD1048576']);
});

test('Previous without a current comment starts at the last one; no comments does nothing', () => {
    const f = commentFixture([{uid: 'a', row: 1, col: 1}, {uid: 'b', row: 4, col: 1}]);
    f.er.gotoComment(-1);
    assert.equal(f.calls.find(c => c[0] === 'select')[1], 'b');
    const none = commentFixture([]);
    none.er.gotoComment(1);
    assert.equal(none.calls.length, 0);
});

// ---- Thesaurus ------------------------------------------------------------------------------------
test('thesaurus meanings are read from the lines of the .dat file', () => {
    const ctx = evalFile(path.join(root, 'apps/common/main/lib/util/Thesaurus.js'));
    const T = ctx.Common.util.Thesaurus;
    const lines = ['other|1', '(noun)|x', 'happy|2', '(adj)|glad|content|joyful', '(adj)|blessed', 'next|1', '(adj)|y'];
    assert.deepEqual(JSON.parse(JSON.stringify(T.parseMeanings(lines, 'happy'))),
        [{pos: 'adj', words: ['glad', 'content', 'joyful']}, {pos: 'adj', words: ['blessed']}]);
    assert.equal(T.parseMeanings(['nothing'], 'happy').length, 0);
});

test('thesaurus files follow the language, with the main variant as fallback', () => {
    const ctx = evalFile(path.join(root, 'apps/common/main/lib/util/Thesaurus.js'));
    ctx.Common.util.LanguageInfo = {getLocalLanguageName: lcid => ({1033: ['en-US', 'English'], 3079: ['de-AT', 'German']}[lcid])};
    const T = ctx.Common.util.Thesaurus;
    assert.equal(T.files(1033)[0], 'en_US/en_US_thes');
    const de = T.files(3079);
    assert.equal(de[0], 'de_AT/de_AT_thes');
    assert.ok(de.includes('de_DE/de_DE_thes'));
});

// ---- PivotTable Analyze: Group Selection, Ungroup, Group Field ----------------------------------
function pivotController(api) {
    const locks = [], events = [];
    const ctx = evalFile(path.join(sse, 'app/controller/PivotTable.js'), {
        Common: {
            enumLock: {noPivot: 'no-pivot', pivotLock: 'pivot-lock', editPivot: 'is-pivot', pivotExpandLock: 'x', pivotCalcItemsLock: 'y', pivotGroupLock: 'pivot-group-lock'},
            Utils: {lockControls: (key, state, opts) => locks.push({key, state, n: opts.array.length})},
            NotificationCenter: {trigger: (...a) => events.push(a)}, util: {}
        }
    });
    const pc = Object.assign({}, ctx.SSE.Controllers.PivotTable);
    pc.api = api;
    pc.view = {btnGroupSelection: {}, btnUngroupPivot: {}, btnGroupField: {}, lockedControls: [], btnsAddPivot: [], btnExpandField: {}, btnCollapseField: {}, btnCalculatedItems: {}};
    pc.appConfig = {isEdit: true};
    return {pc, locks, events};
}

test('Group Selection and Group Field call the engine, Ungroup too', () => {
    const calls = [];
    const f = pivotController({asc_groupPivot: () => calls.push('group'), asc_ungroupPivot: () => calls.push('ungroup')});
    f.pc.onPivotGroup('selection');
    f.pc.onPivotGroup('field');
    f.pc.onPivotGroup('ungroup');
    assert.deepEqual(calls, ['group', 'group', 'ungroup']);
    assert.equal(f.events.length, 3);
});

test('the group buttons follow what the selection allows', () => {
    for (const [canGroup, locked] of [[true, false], [false, true]]) {
        const f = pivotController({asc_canGroupPivot: () => canGroup});
        const info = {asc_getPivotTableInfo: () => ({asc_canExpandCollapseByActiveCell: () => true, asc_canChangeCalculatedItemByActiveCell: () => true}),
            asc_getLockedPivotTable: () => false};
        f.pc.ChangeSettings = () => {};
        f.pc.onSelectionChanged(info);
        const lock = f.locks.find(l => l.key === 'pivot-group-lock');
        assert.equal(lock.state, locked);
        assert.equal(lock.n, 3);
    }
    // outside a pivot table everything is locked, and an engine error never breaks the ribbon
    const f = pivotController({asc_canGroupPivot() { throw new Error('boom'); }});
    assert.equal(f.pc.canGroup(), false);
});

test('the group buttons honour the pivot group lock and have their strings', () => {
    const view = fs.readFileSync(path.join(sse, 'app/view/PivotTable.js'), 'utf8');
    for (const b of ['btnGroupSelection', 'btnUngroupPivot', 'btnGroupField']) {
        const block = view.slice(view.indexOf('this.' + b + ' = new Common.UI.Button'));
        assert.match(block.slice(0, 700), /_set\.pivotGroupLock/, b + ' lock');
    }
    assert.match(fs.readFileSync(path.join(sse, 'app/view/Toolbar.js'), 'utf8'), /pivotGroupLock: 'pivot-group-lock'/);
});

// ---- strings, registration -------------------------------------------------------------------------
const LOCALES = ['en', 'fr', 'de', 'es', 'it', 'ru', 'nl', 'pt', 'pt-pt'];
const KEYS = ['txtGroupSelection', 'txtUngroupPivot', 'txtGroupField', 'tipGroupSelection', 'tipUngroupPivot', 'tipGroupField']
    .map(k => 'SSE.Views.PivotTable.' + k)
    .concat(['txtSpelling', 'tipSpelling', 'txtThesaurus', 'tipThesaurus', 'txtThesaurusTitle', 'txtSearchFor', 'txtLookUp', 'txtSearching',
        'txtNoThesaurus', 'txtNoResults', 'txtTypeWord', 'txtWorkbookStats', 'tipWorkbookStats', 'txtWorkbookStatsTitle', 'txtStatistics',
        'txtStatSheets', 'txtStatHidden', 'txtStatCells', 'txtStatFormulas', 'txtStatTables', 'txtStatPivots', 'txtStatCharts', 'txtStatPictures',
        'txtStatShapes', 'txtStatComments', 'txtNewComment', 'txtPrevComment', 'txtNextComment', 'tipPrevComment', 'tipNextComment',
        'txtShowComments', 'tipShowComments', 'txtGroupPivotgroup'].map(k => 'SSE.Views.RibbonTabs.' + k));

test('Review and PivotTable strings exist in the 9 shipped locales and are translated', () => {
    const en = JSON.parse(fs.readFileSync(path.join(sse, 'locale/en.json'), 'utf8'));
    for (const l of LOCALES) {
        const d = JSON.parse(fs.readFileSync(path.join(sse, 'locale', l + '.json'), 'utf8'));
        for (const k of KEYS) assert.ok(d[k], l + ' misses ' + k);
        if (l !== 'en') {
            assert.notEqual(d['SSE.Views.PivotTable.txtGroupField'], en['SSE.Views.PivotTable.txtGroupField'], l);
            assert.notEqual(d['SSE.Views.RibbonTabs.txtWorkbookStats'], en['SSE.Views.RibbonTabs.txtWorkbookStats'], l);
        }
    }
    const fr = JSON.parse(fs.readFileSync(path.join(sse, 'locale/fr.json'), 'utf8'));
    assert.equal(fr['SSE.Views.PivotTable.txtGroupField'], 'Grouper le champ');
    assert.equal(fr['SSE.Views.RibbonTabs.txtWorkbookStats'], 'Statistiques du classeur');
});

test('every string of the views has its default in the code and its key in en.json', () => {
    const en = JSON.parse(fs.readFileSync(path.join(sse, 'locale/en.json'), 'utf8'));
    const views = {'SSE.Views.RibbonTabs': 'app/view/RibbonTabs.js', 'SSE.Views.PivotTable': 'app/view/PivotTable.js'};
    for (const k of KEYS) {
        const [, , view, name] = /^(SSE\.Views\.\w+)\.(\w+)$/.exec(k) ? [0, 0, k.slice(0, k.lastIndexOf('.')), k.slice(k.lastIndexOf('.') + 1)] : [];
        const src = fs.readFileSync(path.join(sse, views[view]), 'utf8');
        const m = new RegExp('^\\s+' + name + ": '(.*)',?$", 'm').exec(src);
        assert.ok(m, k + ' has no default in the view');
        assert.equal(en[k], m[1].replace(/\\'/g, "'"), k + ' differs from en.json');
    }
});

test('the Review controller and the thesaurus are registered in both app entry points', () => {
    for (const f of ['app.js', 'app_dev.js']) {
        const s = fs.readFileSync(path.join(sse, f), 'utf8');
        assert.match(s, /,'ExcelReview'/);
        assert.match(s, /spreadsheeteditor\/main\/app\/controller\/ExcelReview/);
    }
    assert.match(fs.readFileSync(path.join(sse, 'app/controller/ExcelReview.js'), 'utf8'), /common\/main\/lib\/util\/Thesaurus/);
});

test('Review has Proofing and comment controls in the order of Excel, key tips included', () => {
    const rt = fs.readFileSync(path.join(sse, 'app/controller/RibbonTabs.js'), 'utf8');
    const order = ['slot-elx-rev-spelling', 'slot-elx-rev-thesaurus', 'slot-elx-rev-wbstats', 'slot-elx-rev-prevcomment', 'slot-elx-rev-nextcomment', 'slot-elx-rev-showcomments'];
    let at = -1;
    for (const id of order) {
        const i = rt.indexOf("'" + id + "'");
        assert.ok(i > at, id + ' out of order');
        at = i;
    }
    const tips = fs.readFileSync(path.join(sse, 'app/controller/ExcelKeyTips.js'), 'utf8');
    const letters = order.map(id => new RegExp("S\\('" + id + "'\\), '(\\w+)'").exec(tips)[1]);
    assert.equal(new Set(letters.concat(['C', 'D', 'PS', 'PW', 'U'])).size, order.length + 5, 'review key tips must be unique');
    assert.ok(!rt.includes('arrangeReview: function') || rt.indexOf('me.arrangeReview();') < rt.indexOf('me.mergeProtection();'));
});

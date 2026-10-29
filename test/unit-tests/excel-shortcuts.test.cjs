// SPDX-License-Identifier: AGPL-3.0-only
'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const root = path.resolve(__dirname, '../..');

function load() {
    const notifications = [];
    const context = {
        console,
        SSE: {Controllers: {}},
        Backbone: {Controller: {extend: specification => specification}},
        Asc: {
            c_oAscFillType: {fillDown: 12, fillRight: 13},
            Range: function (c1, r1, c2, r2) { Object.assign(this, {c1, r1, c2, r2}); }
        },
        Common: {NotificationCenter: {trigger: (...values) => notifications.push(values)}},
        define(_dependencies, factory) { factory(); }
    };
    vm.runInNewContext(fs.readFileSync(path.join(root,
        'apps/spreadsheeteditor/main/app/controller/ExcelShortcuts.js'), 'utf8'), context);
    return {controller: context.SSE.Controllers.ExcelShortcuts, Asc: context.Asc, notifications};
}

function fixture(range) {
    const loaded = load(), selections = [], fills = [];
    const worksheet = {
        model: {selectionRange: {ranges: [{...range}], activeCell: {row: range.r1, col: range.c1}}},
        setSelection(value) {
            const copy = {c1: value.c1, r1: value.r1, c2: value.c2, r2: value.r2};
            selections.push(copy);
            this.model.selectionRange = {ranges: [copy], activeCell: {row: copy.r1, col: copy.c1}};
        }
    };
    let selectionChanges = 0;
    const api = {
        wb: {getWorksheet: () => worksheet, _onWSSelectionChanged: () => selectionChanges++},
        asc_FillCells(type) { fills.push({type, range: {...worksheet.model.selectionRange.ranges[0]}}); }
    };
    const instance = {tb: {}, api};
    return {...loaded, worksheet, instance, selections, fills, selectionChanges: () => selectionChanges};
}

test('Fill Down and Fill Right have unique Ctrl/Cmd shortcut bindings', () => {
    const {controller} = load();
    for (const mac of [false, true]) {
        const entries = controller.table.call({mac, mode: {isEdit: true}});
        for (const key of ['KeyD', 'KeyR']) {
            const exact = entries.filter(item => item.key === key && item.c === 1 && item.a === 0 && item.s === 0);
            assert.equal(exact.length, mac ? 2 : 1);
            assert.deepEqual([...exact].map(item => !!item.o.macCtrl).sort(), mac ? [false, true] : [false]);
        }
    }
});

test('single-cell Fill Down copies through the existing engine from the cell above', () => {
    const f = fixture({c1: 2, r1: 4, c2: 2, r2: 4});
    f.controller.fill.call(f.instance, f.Asc.c_oAscFillType.fillDown);
    assert.deepEqual(f.fills, [{type: 12, range: {c1: 2, r1: 3, c2: 2, r2: 4}}]);
    assert.deepEqual(f.selections, [
        {c1: 2, r1: 3, c2: 2, r2: 4},
        {c1: 2, r1: 4, c2: 2, r2: 4}
    ]);
    assert.equal(f.selectionChanges(), 1);
    assert.equal(f.notifications.length, 1);
});

test('single-cell Fill Right copies through the existing engine from the cell to the left', () => {
    const f = fixture({c1: 3, r1: 5, c2: 3, r2: 5});
    f.controller.fill.call(f.instance, f.Asc.c_oAscFillType.fillRight);
    assert.deepEqual(f.fills, [{type: 13, range: {c1: 2, r1: 5, c2: 3, r2: 5}}]);
    assert.deepEqual(f.selections.at(-1), {c1: 3, r1: 5, c2: 3, r2: 5});
});

test('multi-cell fills keep the original selection and first-row/column engine semantics', () => {
    const down = fixture({c1: 0, r1: 0, c2: 1, r2: 3});
    down.controller.fill.call(down.instance, down.Asc.c_oAscFillType.fillDown);
    assert.deepEqual(down.fills, [{type: 12, range: {c1: 0, r1: 0, c2: 1, r2: 3}}]);
    assert.deepEqual(down.selections, []);
    const right = fixture({c1: 0, r1: 0, c2: 3, r2: 2});
    right.controller.fill.call(right.instance, right.Asc.c_oAscFillType.fillRight);
    assert.deepEqual(right.fills, [{type: 13, range: {c1: 0, r1: 0, c2: 3, r2: 2}}]);
    assert.deepEqual(right.selections, []);
});

test('top and left boundary single cells are consumed without an invalid fill', () => {
    const top = fixture({c1: 2, r1: 0, c2: 2, r2: 0});
    top.controller.fill.call(top.instance, top.Asc.c_oAscFillType.fillDown);
    assert.deepEqual(top.fills, []);
    const left = fixture({c1: 0, r1: 2, c2: 0, r2: 2});
    left.controller.fill.call(left.instance, left.Asc.c_oAscFillType.fillRight);
    assert.deepEqual(left.fills, []);
});

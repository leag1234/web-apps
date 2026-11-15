// SPDX-FileCopyrightText: 2026 ElixirOffice contributors
// SPDX-License-Identifier: AGPL-3.0-only
"use strict";

const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const main = path.resolve(__dirname, "../../apps/documenteditor/main");

// a tiny stand-in for the jQuery objects of the lower pane; the caret is the only child
function fakeCaret(pane) {
    const c = {cssv: {}, css(o) { Object.assign(c.cssv, o); return c; }, appendTo() { pane.kids.push(c); return c; }};
    return c;
}
function fakePane() {
    const pane = {kids: [], top: 0};
    pane.find = () => pane;
    pane.children = () => ({length: pane.kids.length, remove() { pane.kids = []; }, css(o) { return pane.kids[0].css(o); }});
    pane.scrollTop = v => { if (v === undefined) return pane.top; pane.top = v; return pane; };
    pane.height = () => 100;
    return pane;
}

function load() {
    const locks = [];
    const context = {
        DE: {Controllers: {}},
        Backbone: {Controller: {extend: props => { const C = function () {}; Object.assign(C.prototype, props); return C; }}},
        _: {extend: Object.assign, bind: (f, o) => f.bind(o), debounce: f => f, throttle: f => f},
        Common: {NotificationCenter: {on() {}, trigger() {}}, Utils: {isMac: false, ModalWindow: {isVisible: () => false}}, UI: {}},
        window: {},
        document: {addEventListener() {}},
        $: () => (context.pane ? fakeCaret(context.pane) : {}),
        console
    };
    context.window.window = context.window;
    context.define = (_d, factory) => factory();
    vm.runInNewContext(fs.readFileSync(path.join(main, "app/controller/WordViews.js"), "utf8"), context, {filename: "WordViews.js"});
    const ctl = new context.DE.Controllers.WordViews();
    ctl._state = {split: {ratio: 0.5, lower: false, target: null}};
    ctl.api = {asc_LockScrollToTarget: v => locks.push(v)};
    return {ctl, locks, context};
}

test("the lower pane takes the keyboard without scrolling the upper pane, and gives it back", () => {
    const {ctl, locks} = load();
    ctl.setSplitLower(true);
    ctl.setSplitLower(true);
    assert.deepEqual(locks, [true]);
    ctl.setSplitLower(false);
    assert.deepEqual(locks, [true, false]);
    assert.equal(ctl._state.split.lower, false);
});

test("engine caret events are remembered and drawn only while the lower pane is active", () => {
    const {ctl} = load();
    const drawn = [];
    ctl.showSplitCaret = reveal => drawn.push(reveal);
    ctl.onSplitTarget(10, 20, 1, 5);
    assert.deepEqual(JSON.parse(JSON.stringify(ctl._state.split.target)), {x: 10, y: 20, page: 1, size: 5});
    assert.deepEqual(drawn, []);
    ctl._state.split.lower = true;
    ctl.onSplitTarget(11, 21, 1, 5);
    assert.deepEqual(drawn, [true]);
});

test("the caret of the lower pane is placed from millimetres and the pane scrolls to it", () => {
    const {ctl, context} = load();
    const pane = context.pane = fakePane();
    ctl.$splitPane = pane;
    ctl.splitGeometry = () => [{w: 200, h: 300, top: 400, left: 10, mmW: 100, mmH: 150}];
    ctl._state.split.lower = true;
    ctl._state.split.target = {x: 10, y: 50, page: 0, size: 5};
    ctl.showSplitCaret(true);
    const caret = pane.kids[0];
    assert.deepEqual(caret.cssv, {left: "30px", top: "500px", height: "10px"});
    assert.equal(pane.top, 500 + 10 - 100 + 20);
    ctl._state.split.lower = false;
    ctl.showSplitCaret(true);
    assert.equal(pane.kids.length, 0);
});

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
        console, setTimeout, clearTimeout, setInterval, clearInterval
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

const BOXES = [
    {w: 200, h: 300, top: 10, left: 10, mmW: 100, mmH: 150},
    {w: 200, h: 300, top: 320, left: 10, mmW: 100, mmH: 150}
];

test("a point of the screen is placed on a page in millimetres, also above and below the pages", () => {
    const {ctl} = load();
    const origin = {left: 50, top: 200};
    // inside the second page: 40 px right of its left edge, 30 px below its top
    assert.deepEqual(JSON.parse(JSON.stringify(ctl.splitLocate(BOXES, origin, 50 + 10 + 40, 200 + 320 + 30))), {page: 1, x: 20, y: 15});
    // above the first page and below the last one the point belongs to those pages, clamped
    const above = ctl.splitLocate(BOXES, origin, 50 + 10 + 40, 200 - 50);
    assert.equal(above.page, 0);
    assert.equal(above.y, 0);
    const below = ctl.splitLocate(BOXES, origin, 50 + 10 + 40, 200 + 5000);
    assert.equal(below.page, 1);
    assert.equal(below.y, 150);
    assert.equal(ctl.splitLocate([], origin, 0, 0), null);
});

function clickSetup() {
    const {ctl, context} = load();
    const calls = [];
    ctl.$splitPane = {find: () => ({offset: () => ({left: 50, top: 200}), children: () => ({length: 0, remove() {}}), remove() {}})};
    ctl.splitGeometry = () => BOXES;
    ctl.scheduleSplitSelection = () => {};
    context.document.getElementById = () => null;
    ctl.startSplitDrag = (count, e) => calls.push(["drag", count, e.pageX]);
    ctl.api = {
        asc_LockScrollToTarget() {},
        asc_enableKeyEvents() {},
        asc_MouseInPage: (...a) => calls.push(a)
    };
    return {ctl, calls, context};
}
const $page = {attr: () => "0"};

test("a click, a double click and a triple click of the lower pane reach the engine with their click count", () => {
    const {ctl, calls} = clickSetup();
    [1, 2, 3, 4].forEach(detail => ctl.onSplitPageClick(
        {which: 1, originalEvent: {detail}, pageX: 160, pageY: 260, shiftKey: false, ctrlKey: false, preventDefault() {}}, $page));
    const downs = calls.filter(c => c[0] === "down");
    assert.deepEqual(downs.map(c => c[4]), [1, 2, 3, 3]);
    assert.deepEqual(downs[0].slice(1, 4), [0, 50, 25]);
    assert.equal(calls.filter(c => c[0] === "drag").length, 4);
    assert.equal(ctl._state.split.lower, true);
});

test("only the left button and the shift key are passed on", () => {
    const {ctl, calls} = clickSetup();
    ctl.onSplitPageClick({which: 3, pageX: 160, pageY: 260, preventDefault() {}}, $page);
    assert.equal(calls.length, 0);
    ctl.onSplitPageClick({which: 1, pageX: 160, pageY: 260, shiftKey: true, ctrlKey: false, preventDefault() {}}, $page);
    const down = calls.find(c => c[0] === "down");
    assert.equal(down[5], true);
});

test("dragging with the button down moves the selection until the button is released", () => {
    const {ctl, calls, context} = clickSetup();
    const handlers = {}, off = [];
    context.$ = () => ({
        on(name, fn) { handlers[name] = fn; return this; },
        off(name) { off.push(name); return this; }
    });
    delete ctl.startSplitDrag;
    ctl.$splitPane.offset = () => ({top: 100});
    ctl.$splitPane.height = () => 400;
    ctl.$splitPane.scrollTop = () => 0;
    const ev = (x, y) => ({pageX: x, pageY: y, shiftKey: false, ctrlKey: false, metaKey: false});
    ctl.startSplitDrag(2, ev(160, 260));
    assert.ok(ctl._state.split.drag);
    handlers["mousemove.elxsplitdrag"](ev(200, 300));
    handlers["mouseup.elxsplitdrag"](ev(240, 330));
    const kinds = calls.map(c => c[0] + ":" + c[4]);
    assert.deepEqual(kinds, ["move:2", "up:2"]);
    assert.equal(ctl._state.split.drag, null);
    assert.ok(off.length >= 1);
});

test("the selection of the document is painted on the pages of the lower pane", () => {
    const {ctl, context} = load();
    const appended = [];
    const kids = {remove() {}};
    const pages = {children: () => kids, append: html => appended.push(html)};
    ctl.$splitPane = {find: () => pages, scrollTop: () => 0, height: () => 1000};
    ctl.splitGeometry = () => BOXES;
    ctl.api = {asc_GetSelectionRects: i => i === 0 ? [{x: 10, y: 20, w: 30, h: 5}] : []};
    ctl.showSplitSelection();
    assert.equal(appended.length, 1);
    assert.match(appended[0], /left:30px;top:50px;width:60px;height:10px/);
    assert.match(appended[0], /opacity:0\.3/);
    ctl.api.asc_GetSelectionRects = () => [];
    ctl.showSplitSelection();
    assert.equal(appended.length, 1, "no layer without a selection");
});

// SPDX-FileCopyrightText: 2026 ElixirOffice contributors
// SPDX-License-Identifier: AGPL-3.0-only
'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const root = path.resolve(__dirname, '../..');

function el(tag) {
    return {
        tagName: tag, children: [], attrs: {}, dataset: {}, style: {},
        width: 1, height: 1, offsetTop: 0, offsetHeight: 0,
        getAttribute(name) { return this.attrs[name]; },
        scrollIntoView() { this.scrolled = true; }
    };
}

function wrap(node) {
    const list = Array.isArray(node) ? node : [node];
    return {
        0: list[0], length: list.length,
        find(sel) {
            if (sel === '.elx-pages-list') return wrap(list[0].list);
            const m = sel.match(/^\.elx-page-thumb\[data-page=(\d+)\] canvas$/);
            if (m) {
                const b = list[0].buttons.find(x => x.attrs['data-page'] === m[1]);
                return wrap(b ? [b.children[0]] : []);
            }
            if (sel === '.elx-page-thumb') return wrap(list[0].buttons);
            if (sel === '.elx-page-thumb canvas') return wrap(list[0].buttons.map(b => b.children[0]));
            return wrap([]);
        },
        html(markup) {
            const re = /<button[^>]*data-page="(\d+)"[^>]*><canvas width="1" height="1" style="aspect-ratio: 150 \/ (\d+);"><\/canvas><span>\d+<\/span><\/button>/g;
            list[0].buttons = [];
            let m;
            while ((m = re.exec(markup))) {
                const b = el('button');
                b.attrs['data-page'] = m[1];
                const c = el('canvas');
                c.style['aspect-ratio'] = '150 / ' + m[2];
                b.children.push(c);
                list[0].buttons.push(b);
            }
        },
        each(fn) { list.forEach((b, i) => fn.call(b, i, b)); },
        get() { return list.slice(); },
        is() { return true; }
    };
}

function loadView(observerImpl) {
    let view;
    const pagesEl = el('div');
    pagesEl.list = el('div');
    pagesEl.scrollTop = 0;
    pagesEl.clientHeight = 800;
    const $pages = wrap(pagesEl);
    const drawn = [];
    const api = {
        getCountPages: () => 300,
        getCurrentPage: () => 0,
        asc_getPageSize: () => ({ W: 210, H: 297 })
    };
    const context = {
        console, setTimeout, clearTimeout, Set, Math, parseInt, window: {},
        define(deps, factory) { factory(); view = context.DE.Views.ElixirNavPane; },
        Backbone: { View: { extend: o => o } },
        _: { extend: (a, b) => Object.assign(a, b) },
        $: Object.assign(function () { return wrap([]); }, {}),
        Common: { Utils: { String: {
            htmlEncode: s => s,
            format: (s, v) => s.replace('{0}', v)
        } } },
        DE: { Views: {}, getController: () => ({ api }) }
    };
    context.window.IntersectionObserver = observerImpl;
    context.IntersectionObserver = observerImpl;
    context.define.amd = true;
    vm.runInNewContext(fs.readFileSync(path.join(root,
        'apps/documenteditor/main/app/elixir-chrome/NavPane.js'), 'utf8'), context);
    view.api = api;
    view.pageCount = 300;
    view.$pages = $pages;
    view.drawThumb = function (i, canvas) { drawn.push(i); canvas.width = 150; canvas.height = 212; };
    return { view, drawn, pagesEl };
}

const sleep = ms => new Promise(r => setTimeout(r, ms));

test('only thumbs near the visible area are drawn, far ones are released', async () => {
    let ioCallback;
    const observed = [];
    class IO {
        constructor(cb) { ioCallback = cb; }
        observe(t) { observed.push(t); }
        disconnect() { this.disconnected = true; }
    }
    const { view, drawn, pagesEl } = loadView(IO);
    view.renderPages();
    assert.equal(observed.length, 300);
    assert.equal(pagesEl.list.buttons.length, 300);
    assert.equal(pagesEl.list.buttons[0].children[0].style['aspect-ratio'], '150 / 212');
    // pages 0..9 visible (within rootMargin), page 50 not
    ioCallback([0, 1, 2, 3, 4, 5, 6, 7, 8, 9].map(p => ({
        target: pagesEl.list.buttons[p], isIntersecting: true
    })));
    await sleep(150);
    assert.deepEqual([...drawn].sort((a, b) => a - b), [0, 1, 2, 3, 4, 5, 6, 7, 8, 9]);
    assert.equal(pagesEl.list.buttons[50].children[0].dataset.drawn, undefined);
    // page 3 scrolls far away: evicted and released
    ioCallback([{ target: pagesEl.list.buttons[3], isIntersecting: false }]);
    const c3 = pagesEl.list.buttons[3].children[0];
    assert.equal(c3.dataset.drawn, undefined);
    assert.equal(c3.width, 1);
    assert.equal(c3.height, 1);
    assert.equal(c3.style['aspect-ratio'], '150 / 212');
    // re-render disconnects the old observer
    const oldObserver = view._pagesObserver;
    view.renderPages();
    assert.equal(oldObserver.disconnected, true);
    clearTimeout(view._pagesTimer);
});

test('falls back to drawing everything without IntersectionObserver', async () => {
    const { view, drawn } = loadView(undefined);
    view.renderPages();
    await sleep(1600);
    assert.equal(drawn.length, 300);
    clearTimeout(view._pagesTimer);
});

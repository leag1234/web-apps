// SPDX-FileCopyrightText: 2026 ElixirOffice contributors
// SPDX-License-Identifier: AGPL-3.0-only
'use strict';
// Issue 65: the packaged Open Sans (core-fonts, Open Sans 1.10) is registered as "Open Sans Semibold" and
// "Open Sans Extrabold" (name ID 1 of those files); documents written with Open Sans 3.x name the same faces
// "Open Sans SemiBold" / "Open Sans ExtraBold". Word matches font names without regard to case. The engine
// map is case-sensitive, so the packaged 0.7.2 app drew those runs with Carlito. The names below are the real ones.

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const root = path.resolve(__dirname, '../..');

function openEngine() {
    const asked = [];
    const families = ['Carlito', 'Open Sans', 'Open Sans Extrabold', 'Open Sans Light', 'Open Sans Semibold'];
    const map = {};
    const infos = families.map((name, i) => { map[name] = i; return {Name: name}; });
    const app = {
        DocumentEmbedded: {},
        // the SDK answers a name it does not have with the unknown-font substitute
        GetFontFileWeb(name) { asked.push(name); return {m_wsFontName: map[name] !== undefined ? name : 'Carlito'}; },
        GetFontFile(name, style) { asked.push(name + '/' + style); return {m_wsFontName: map[name] !== undefined ? name : 'Carlito'}; }
    };
    const context = {
        ArrayBuffer, Blob, Uint8Array, console, setTimeout, clearTimeout,
        clearInterval() {}, setInterval: () => 0,
        performance: {now: () => 0},
        localStorage: {getItem: () => null, setItem() {}, removeItem() {}}
    };
    context.window = context;
    context.Common = {};
    context.AscFonts = {
        CFont: function (name) { this.name = name; },
        CFontFileLoader: function (id) { this.Id = id; this.Status = -1; },
        CFontInfo: function (name) { this.Name = name; },
        g_fontApplication: app, g_font_files: [], g_font_infos: infos, g_map_font_index: map
    };
    context.AscCommon = {g_font_loader: {LoadDocumentFonts() {}, LoadFonts(names, callback) { callback(); }}};
    context.define = (_dependencies, factory) => factory();
    vm.runInNewContext(fs.readFileSync(path.join(root, 'apps/common/main/lib/elixir-localfonts/Core.js'), 'utf8'), context, {filename: 'Core.js'});
    const LF = context.Common.ElixirLocalFonts;
    assert.equal(LF.install(), true);
    return {app, asked, LF, context};
}

test('a document font name differing only in case reaches the packaged family', () => {
    const {app, asked} = openEngine();
    assert.equal(app.GetFontFileWeb('Open Sans SemiBold').m_wsFontName, 'Open Sans Semibold');
    assert.equal(app.GetFontFileWeb('Open Sans ExtraBold').m_wsFontName, 'Open Sans Extrabold');
    assert.equal(app.GetFontFile('Open Sans SemiBold', 1).m_wsFontName, 'Open Sans Semibold');
    // the style argument is kept
    assert.equal(asked.at(-1), 'Open Sans Semibold/1');
});

test('exact, unknown and embedded names are left to the SDK unchanged', () => {
    const {app, asked} = openEngine();
    assert.equal(app.GetFontFileWeb('Open Sans').m_wsFontName, 'Open Sans');
    assert.equal(app.GetFontFileWeb('Open Sans Bold Test').m_wsFontName, 'Carlito');
    assert.deepEqual(asked, ['Open Sans', 'Open Sans Bold Test']);
    app.DocumentEmbedded['open sans semibold'] = {m_wsFontName: 'embedded'};
    app.GetFontFileWeb('open sans semibold');
    assert.equal(asked.at(-1), 'open sans semibold', 'a font embedded in the document keeps its own name');
});

test('names registered later are matched too', () => {
    const {app, context} = openEngine();
    const A = context.AscFonts;
    assert.equal(app.GetFontFileWeb('NOTO SERIF').m_wsFontName, 'Carlito');
    A.g_map_font_index['Noto Serif'] = A.g_font_infos.length;
    A.g_font_infos.push({Name: 'Noto Serif'});
    assert.equal(app.GetFontFileWeb('NOTO SERIF').m_wsFontName, 'Noto Serif');
});

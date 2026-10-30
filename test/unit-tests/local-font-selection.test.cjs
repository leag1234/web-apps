// SPDX-FileCopyrightText: 2026 ElixirOffice contributors
// SPDX-License-Identifier: AGPL-3.0-only
'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const root = path.resolve(__dirname, '../..');

function loadCore(succeed = true) {
    const files = [];
    const infos = [];
    const loaded = [];
    const context = {
        ArrayBuffer,
        Blob,
        Uint8Array,
        clearInterval() {},
        clearTimeout,
        console,
        performance: {now: () => 0},
        setInterval: () => 0,
        setTimeout,
        localStorage: {getItem: () => null, setItem() {}, removeItem() {}}
    };
    context.window = context;
    context.Common = {};
    context.AscFonts = {
        CFont: function (name) { this.name = name; },
        CFontFileLoader: function (id) { this.Id = id; this.Status = -1; },
        CFontInfo: function (name, _unused, r, _rFace, i, _iFace, b, _bFace, bi) {
            this.Name = name;
            this.indexR = r;
            this.indexI = i;
            this.indexB = b;
            this.indexBI = bi;
        },
        g_fontApplication: {},
        g_font_files: files,
        g_font_infos: infos,
        g_map_font_index: {}
    };
    context.AscCommon = {
        g_font_loader: {
            LoadDocumentFonts() {},
            LoadFonts(names, callback) {
                loaded.push(names.slice());
                for (const name of names) {
                    const info = infos[context.AscFonts.g_map_font_index[name]];
                    for (const key of ['indexR', 'indexI', 'indexB', 'indexBI']) {
                        if (info[key] >= 0) files[info[key]].Status = succeed ? 0 : 1;
                    }
                }
                callback();
            }
        }
    };
    context.define = (_dependencies, factory) => factory();
    vm.runInNewContext(
        fs.readFileSync(path.join(root, 'apps/common/main/lib/elixir-localfonts/Core.js'), 'utf8'),
        context,
        {filename: 'Core.js'}
    );
    const local = context.Common.ElixirLocalFonts;
    local.fonts = [{
        family: 'Fira Sans', fullName: 'Fira Sans', postscriptName: 'FiraSans-Regular',
        style: 'Regular', blob: async () => new Blob([])
    }];
    local.register();
    return {local, loaded, files};
}

test('known local font loads before selection continues', () => {
    const {local, loaded} = loadCore();
    let ready;
    assert.equal(local.loadForSelection('Fira Sans', value => { ready = value; }), true);
    assert.deepEqual(JSON.parse(JSON.stringify(loaded)), [['Fira Sans']]);
    assert.equal(ready, true);
});

test('unknown font is left to the existing unavailable-font path', () => {
    const {local, loaded} = loadCore();
    let called = false;
    assert.equal(local.loadForSelection('Not Installed', () => { called = true; }), false);
    assert.deepEqual(loaded, []);
    assert.equal(called, false);
});

test('unreadable local font returns to the unavailable-font warning path', () => {
    const {local} = loadCore(false);
    let ready;
    assert.equal(local.loadForSelection('Fira Sans', value => { ready = value; }), true);
    assert.equal(ready, false);
});

test('every editor font-selection entry point uses the local load gate', () => {
    for (const file of [
        'apps/documenteditor/main/app/controller/Toolbar.js',
        'apps/spreadsheeteditor/main/app/controller/Toolbar.js',
        'apps/presentationeditor/main/app/controller/Toolbar.js',
        'apps/pdfeditor/main/app/controller/Toolbar.js',
        'apps/pdfeditor/main/app/controller/DocumentHolderExt.js'
    ]) assert.match(fs.readFileSync(path.join(root, file), 'utf8'),
        /ElixirLocalFonts\.loadForSelection/, file);
});

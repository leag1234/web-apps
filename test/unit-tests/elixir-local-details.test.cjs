// SPDX-FileCopyrightText: 2026 ElixirOffice contributors
// SPDX-License-Identifier: AGPL-3.0-only
'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const root = path.resolve(__dirname, '../..');
const localeRoot = path.join(root, 'apps/documenteditor/main/locale');
const localeKeys = [
    'DE.Views.ElixirLayoutOptions.textBehind',
    'DE.Views.ElixirLayoutOptions.textClose',
    'DE.Views.ElixirLayoutOptions.textFix',
    'DE.Views.ElixirLayoutOptions.textInFront',
    'DE.Views.ElixirLayoutOptions.textInline',
    'DE.Views.ElixirLayoutOptions.textInlineGroup',
    'DE.Views.ElixirLayoutOptions.textMove',
    'DE.Views.ElixirLayoutOptions.textSeeMore',
    'DE.Views.ElixirLayoutOptions.textSquare',
    'DE.Views.ElixirLayoutOptions.textThrough',
    'DE.Views.ElixirLayoutOptions.textTight',
    'DE.Views.ElixirLayoutOptions.textTitle',
    'DE.Views.ElixirLayoutOptions.textTopBottom',
    'DE.Views.ElixirLayoutOptions.textWrapGroup',
    'DE.Views.ElixirLayoutOptions.tipAnchor',
    'DE.Views.ElixirLayoutOptions.tipLockedAnchor',
    'DE.Views.HeaderFooterTab.capBtnInsImage',
    'DE.Views.HeaderFooterTab.mniImageFromFile',
    'DE.Views.HeaderFooterTab.mniImageFromStorage',
    'DE.Views.HeaderFooterTab.mniImageFromUrl',
    'DE.Views.HeaderFooterTab.tipInsertImage',
    'DE.Views.RibbonTabs.txtPictures'
];

function loadAmd(relative, context) {
    let value;
    context.define = function (_dependencies, factory) {
        value = factory(...(context.amdArguments || []));
    };
    context.define.amd = true;
    vm.runInNewContext(fs.readFileSync(path.join(root, relative), 'utf8'), context, {
        filename: relative
    });
    return value;
}

test('all document editor locales contain the Z17 picture and layout strings', () => {
    const files = fs.readdirSync(localeRoot).filter(name => name.endsWith('.json')).sort();
    assert.equal(files.length, 45);
    for (const file of files) {
        const strings = JSON.parse(fs.readFileSync(path.join(localeRoot, file), 'utf8'));
        for (const key of localeKeys) {
            assert.equal(typeof strings[key], 'string', file + ': ' + key);
            assert.notEqual(strings[key].trim(), '', file + ': ' + key);
        }
    }
});

test('font guidance has a complete non-empty table for every editor locale', () => {
    const context = {Common: {}};
    context.window = context;
    const strings = loadAmd('apps/common/main/lib/elixir-localfonts/Strings.js', context);
    const languages = Object.keys(strings).sort();
    const localeLanguages = fs.readdirSync(localeRoot).filter(name => name.endsWith('.json'))
        .map(name => name.slice(0, -5)).sort();
    assert.deepEqual(languages, localeLanguages);
    const keys = Object.keys(strings.en).sort();
    for (const language of languages) {
        assert.deepEqual(Object.keys(strings[language]).sort(), keys, language);
        for (const key of keys) {
            assert.equal(typeof strings[language][key], 'string', language + ': ' + key);
            assert.notEqual(strings[language][key].trim(), '', language + ': ' + key);
        }
    }
});

test('font summaries isolate every name and truncate only after four entries', () => {
    const context = {Common: {Controllers: {}}};
    context.window = context;
    loadAmd('apps/common/main/lib/elixir-localfonts/Strings.js', context);
    context.Backbone = {
        Controller: {
            extend(specification) {
                function Controller() {}
                Controller.prototype = specification;
                return Controller;
            }
        }
    };
    const Controller = loadAmd('apps/common/main/lib/elixir-localfonts/Controller.js', context);
    const namesText = Controller.prototype.namesText;
    const names = ['Arial', String.fromCodePoint(0x05e2, 0x05d1, 0x05e8, 0x05d9, 0x05ea),
        String.fromCodePoint(0x0646, 0x0633, 0x062e), 'Latin', 'Fifth'];
    const isolate = name => String.fromCodePoint(0x2068) + name + String.fromCodePoint(0x2069);
    const firstFour = names.slice(0, 4).map(isolate).join(', ');
    assert.equal(namesText(names.slice(0, 4)), firstFour);
    assert.equal(namesText(names), firstFour + String.fromCodePoint(0x2026));
});

test('missing imported offline pictures resolve locally without hiding valid or remote URLs', () => {
    const context = {Common: {}};
    context.window = context;
    context.amdArguments = [{}, {}, {}, {}, {resolve() {}}, {}];
    const local = loadAmd('apps/common/main/lib/elixir-local/Controller.js', context);
    const imported = new Map([['media/real.png', 'blob:real-picture']]);
    assert.equal(local.resolveImportedImage(imported, 'media/real.png', '_offline_media/real.png'),
        'blob:real-picture');
    assert.equal(local.resolveImportedImage(imported, 'missing.png', 'https://example.test/missing.png'),
        'https://example.test/missing.png');
    assert.match(local.resolveImportedImage(imported, 'image-1.jpg',
        'eoapp://desktop.elixiroffice.invalid/example/local-assets/web-apps/apps/documenteditor/main/_offline_media/image-1.jpg'),
        /^data:image\/gif;base64,/);
});

// SPDX-License-Identifier: AGPL-3.0-only
'use strict';
// PDF editor in the local desktop app: the quick access toolbar shows Print although the editor's own print panel is off.
const test = require('node:test'), assert = require('node:assert/strict'), fs = require('node:fs'), path = require('node:path');
const root = path.resolve(__dirname, '../..');
const read = f => fs.readFileSync(path.join(root, f), 'utf8');
const main = read('apps/pdfeditor/main/app/controller/Main.js');
const header = read('apps/common/main/lib/view/Header.js');

test('the local desktop app flags native printing when the editor print panel is off', () => {
    assert.match(main, /eoNativePrint = !this\.appOptions\.canPrint && !!Common\.ElixirLocal;/);
});
test('the quick access Print button is created for native printing too', () => {
    assert.match(header, /\(config\.canPrint \|\| config\.eoNativePrint\) && config\.twoLevelHeader/);
});
test('the Header print click asks the application to print instead of opening the print panel', () => {
    const m = /onPrint: function\(\) \{([\s\S]*?)\n            \},/.exec(main);
    assert.ok(m);
    const body = m[1];
    assert.ok(body.indexOf("trigger('eo:native-print')") !== -1);
    assert.ok(body.indexOf('eoNativePrint') < body.indexOf("trigger('file:print')"), 'native check comes first');
});
test('the Customize quick access menu lists Print when the button exists', () => {
    assert.match(header, /if \(me\.btnPrint\) \{\s*arr\.push\(\{\s*caption: me\.textPrint,\s*value: 'print'/);
});

test('the ribbon print lock does not disable the quick access Print button in native print mode', () => {
    const vp = read('apps/pdfeditor/main/app/controller/Viewport.js');
    const m = /'print:disabled' : function \(state\) \{([\s\S]*?)\n                    \},/.exec(vp);
    assert.ok(m);
    assert.match(m[1], /var nativePrint = me\.getApplication\(\)\.getController\('Main'\)\.appOptions\.eoNativePrint;/);
    assert.match(m[1], /if \( me\.header\.btnPrint && !nativePrint \)\s*me\.header\.btnPrint\.setDisabled\(state\);/);
});

test('the native Print button of the quick access toolbar is created enabled', () => {
    assert.match(header, /createTitleButton\('toolbar__icon icon--inverse btn-print', \$html\.findById\('#slot-btn-dt-print'\), !config\.eoNativePrint,/);
});

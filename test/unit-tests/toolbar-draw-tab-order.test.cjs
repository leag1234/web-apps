'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');

const root = path.resolve(__dirname, '../..');
const controllers = [
    'apps/documenteditor/main/app/controller/Toolbar.js',
    'apps/presentationeditor/main/app/controller/Toolbar.js',
    'apps/spreadsheeteditor/main/app/controller/Toolbar.js'
];

test('the Draw tab is inserted after the Insert tab in every editor', () => {
    for (const controller of controllers) {
        const source = fs.readFileSync(path.join(root, controller), 'utf8');
        const draw = source.indexOf("action: 'draw'");
        assert.ok(draw >= 0, controller + ': Draw tab definition');
        const insertion = source.indexOf("getTabIdx('ins')", draw);
        assert.ok(insertion >= 0, controller + ': Draw tab uses the Insert tab index');
        const addTab = source.indexOf('addTab(tab, $panel, insIdx < 0 ? 2 : insIdx)', insertion);
        assert.ok(addTab >= 0, controller + ': Draw tab keeps the fallback insertion index');
    }
});

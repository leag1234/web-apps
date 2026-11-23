'use strict';

const assert = require('node:assert');
const test = require('node:test');
const fs = require('node:fs');
const path = require('node:path');

const root = path.resolve(__dirname, '../..');
const read = file => fs.readFileSync(path.join(root, file), 'utf8');

// #99: adding an item must not search the whole list again (quadratic time on large AutoFilter lists).
for (const file of ['TreeView.js', 'DataView.js']) {
    test(file + ' reuses the list container while it is in the view (#99)', () => {
        const source = read('apps/common/main/lib/component/' + file);
        const start = source.indexOf('onAddItem: function(record, store, opts)');
        assert.ok(start > 0);
        const body = source.slice(start, start + 1400);
        assert.match(body, /var innerEl = this\._eoInnerEl;\s*\n\s*if \(!innerEl \|\| !innerEl\[0\] \|\| !this\.el\.contains\(innerEl\[0\]\)\)\s*\n\s*innerEl = this\._eoInnerEl = /);
    });
}

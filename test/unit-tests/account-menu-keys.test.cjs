'use strict';

const assert = require('node:assert');
const test = require('node:test');
const fs = require('node:fs');
const path = require('node:path');

const source = fs.readFileSync(path.join(__dirname, '../../apps/common/main/lib/elixir-chrome/TitleBar.js'), 'utf8');

test('the account menu keyboard handler is bound once per element (#100)', () => {
    assert.match(source, /this\.\$account\.off\('keydown\.elxaccount'\)\.on\('keydown\.elxaccount', function \(e\)/);
    assert.doesNotMatch(source, /this\.\$account\.on\('keydown', function/);
});

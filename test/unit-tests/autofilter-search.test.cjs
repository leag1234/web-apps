'use strict';

const assert = require('node:assert');
const test = require('node:test');
const fs = require('node:fs');
const path = require('node:path');

const root = path.resolve(__dirname, '../..');
const read = file => fs.readFileSync(path.join(root, file), 'utf8');
const source = read('apps/spreadsheeteditor/main/app/view/AutoFilterDialog.js');
const match = source.match(/SSE\.Views\.AutoFilterDialog\.eoSearchPattern = function \(value\) \{([\s\S]*?)\n    \};/);
assert(match, 'eoSearchPattern should be exported for tests');
const helperMatch = source.match(/SSE\.Views\.AutoFilterDialog\.isCappedFilterValueList = function \(valueCount\) \{[\s\S]*?SSE\.Views\.AutoFilterDialog\.getSearchCustomFilterValue = function \(searchText\) \{[\s\S]*?\n    \};/);
assert(helperMatch, 'search fallback helpers should be exported for tests');
const SSE = { Views: { AutoFilterDialog: {} } };
new Function('SSE', `${helperMatch[0]}\n${match[0]}`)(SSE);
const { eoSearchPattern, shouldUseSearchCustomFilter, getSearchCustomFilterValue, areAllSearchMatchesChecked, hasSearchWildcard } = SSE.Views.AutoFilterDialog;
const matches = (pattern, value) => new RegExp(eoSearchPattern(pattern), 'i').test(value);

test('Excel wildcards are translated to a case-insensitive RegExp (contains without wildcards, whole value with them)', () => {
    assert.ok(matches('RC*99', 'RC-12399'));
    assert.ok(matches('RC*99', 'RC-99'));
    assert.ok(!matches('RC*99', 'RB-99'));
    assert.ok(matches('1.2*2500*1250*', '1.2*2500*1250*14'));
    assert.ok(matches('1.2*2500*1250*', '1.2x2500y1250'));
    assert.ok(!matches('1.2*2500*1250*', 'v1.2x2500y1250'));
    assert.ok(matches('a?c', 'abc'));
    assert.ok(!matches('a?c', 'ac'));
    assert.ok(matches('5~*2', '5*2'));
    assert.ok(!matches('5~*2', '5x2'));
    assert.ok(matches('(1)', '(1)'));
    assert.ok(matches('[a]', '[a]'));
});

test('the search RegExp has no global flag, avoiding lastIndex state', () => {
    assert.doesNotMatch(source, /new RegExp\([^)]*,\s*'ig'\)/);
    assert.match(source, /new RegExp\(SSE\.Views\.AutoFilterDialog\.eoSearchPattern\(value\), 'i'\)/);
});

test('typing is debounced and Enter rebuilds immediately', () => {
    assert.match(source, /me\.clearSearchTimer\(\);\s*\n\s*me\.searchTimer = _\.delay\(function \(\) \{ me\.searchTimer = undefined; me\.setupDataCells\(\); \}, 250\);/);
    assert.match(source, /event\.keyCode == Common\.UI\.Keys\.RETURN[\s\S]*?me\.clearSearchTimer\(\);\s*\n\s*me\.setupDataCells\(\);/);
});

test('the list is capped at 10,000 values and hidden states are preserved', () => {
    assert.match(source, /visibleValueCount >= 10000/);
    assert.match(source, /textNotAllItems\s*:\s*'Not all items showing'/);
    assert.match(source, /for \(var throughIndex = 2; throughIndex < idxs\.length; throughIndex\+\+\)/);
    assert.match(read('apps/spreadsheeteditor/main/locale/en.json'), /"SSE\.Views\.AutoFilterDialog\.textNotAllItems": "Not all items showing"/);
    assert.match(read('apps/spreadsheeteditor/main/locale/fr.json'), /"SSE\.Views\.AutoFilterDialog\.textNotAllItems": "Tous les éléments ne sont pas affichés"/);
    assert.match(read('apps/spreadsheeteditor/main/locale/de.json'), /"SSE\.Views\.AutoFilterDialog\.textNotAllItems": "Nicht alle Elemente werden angezeigt"/);
});

test('bulk selection during a search never touches values that do not match (#99)', () => {
    assert.match(source, /if \(!me\.filter \|\| \(me\.filterMatches && me\.filterMatches\[throughIndex\]\)\)\s*\n\s*idxs\[throughIndex\] = check;/);
    assert.match(source, /filterMatches\[throughIndex\] = applyfilter;/);
    assert.match(source, /hiddenMatches\.push\(throughIndex\);/);
    assert.match(source, /\(this\.hiddenMatches \|\| \[\]\)\.forEach\(function\(throughIndex\) \{\s*\n\s*arr\[throughIndex-2\]\.asc_setVisible\(!!me\.filteredIndexes\[throughIndex\]\);/);
});

test('the debounce timer helper belongs to AutoFilterDialog, not to another dialog (#99)', () => {
    const start = source.indexOf('SSE.Views.AutoFilterDialog = Common.UI.Window.extend(');
    assert.ok(start > 0);
    const before = source.slice(0, start), inside = source.slice(start);
    assert.doesNotMatch(before, /clearSearchTimer/);
    assert.match(inside, /clearSearchTimer: function \(\) \{\s*\n\s*if \(this\.searchTimer\)/);
});

test('a complete search on a capped list uses a custom filter that follows the typed text (#99)', () => {
    const options = {
        searchText: 'RC*99',
        valueCount: 10000,
        allMatchesChecked: true,
        addSelectionChecked: false,
        hasNonTextSearchMatch: false,
        isPivot: false
    };
    assert.equal(getSearchCustomFilterValue(options.searchText), 'RC*99');
    assert.equal(shouldUseSearchCustomFilter(options), true);
});

test('partial selections, add-to-filter, pivots, and uncapped lists keep value filtering (#99)', () => {
    const options = {
        searchText: 'RC*99',
        valueCount: 10000,
        allMatchesChecked: true,
        addSelectionChecked: false,
        hasNonTextSearchMatch: false,
        isPivot: false
    };
    assert.equal(shouldUseSearchCustomFilter({ ...options, allMatchesChecked: false }), false);
    assert.equal(shouldUseSearchCustomFilter({ ...options, hasNonTextSearchMatch: true }), false);
    assert.equal(shouldUseSearchCustomFilter({ ...options, addSelectionChecked: true }), false);
    assert.equal(shouldUseSearchCustomFilter({ ...options, isPivot: true }), false);
    assert.equal(shouldUseSearchCustomFilter({ ...options, valueCount: 9999 }), false);
    assert.equal(shouldUseSearchCustomFilter({ ...options, searchText: '' }), false);
});

test('the custom filter preserves typed Excel wildcards and escapes (#99)', () => {
    assert.equal(getSearchCustomFilterValue('RC*99'), 'RC*99');
    assert.equal(getSearchCustomFilterValue('5~*2'), '*5~*2*');
    assert.equal(getSearchCustomFilterValue('a?b'), 'a?b');
});

test('plain text searches contain, typed wildcards are never wrapped (#99b)', () => {
    assert.equal(getSearchCustomFilterValue('WOOD'), '*WOOD*');
    assert.equal(getSearchCustomFilterValue('WOOD*'), 'WOOD*');
    assert.equal(getSearchCustomFilterValue('*WOOD'), '*WOOD');
    assert.equal(getSearchCustomFilterValue('W?OD'), 'W?OD');
    assert.equal(hasSearchWildcard('WOOD'), false);
    assert.equal(hasSearchWildcard('5~*2'), false);
    assert.equal(hasSearchWildcard('5~~*2'), true);
    assert.equal(hasSearchWildcard('a?'), true);
    assert.equal(hasSearchWildcard(undefined), false);
});

test('WOOD* in the search box means begins with WOOD, as in Excel (#99b)', () => {
    assert.ok(matches('WOOD*', 'WOOD plank'));
    assert.ok(matches('WOOD*', 'wood'));
    assert.ok(!matches('WOOD*', 'BIGWOOD plank'));
    assert.ok(matches('*WOOD', 'BIG WOOD'));
    assert.ok(!matches('*WOOD', 'WOOD plank'));
    assert.ok(matches('WOOD', 'BIGWOOD plank'));
    assert.ok(matches('W?OD*', 'WOOD plank'));
    assert.ok(!matches('W?OD*', 'WD plank'));
});

test('OK applies a pending debounced search first (#99b)', () => {
    assert.match(source, /save: function \(\) \{\s*\n\s*if \(this\.searchTimer\) \{[^}]*this\.clearSearchTimer\(\);\s*\n\s*this\.setupDataCells\(\);/);
});

test('the capped-list notice remains visible during a search (#99)', () => {
    assert.match(source, /isCappedValueList = SSE\.Views\.AutoFilterDialog\.isCappedFilterValueList\(me\.configTo\.asc_getValues\(\)\.length\)/);
    assert.match(source, /if \(visibleValueCount >= 10000 \|\| \(me\.filter && isCappedValueList\)\)/);
});

test('the fallback is restricted when listed matches include numbers or dates (#99)', () => {
    assert.match(source, /if \(applyfilter && \(isnumber \|\| isDate\)\)\s*\n\s*hasNonTextSearchMatch = true;/);
    assert.match(source, /hasNonTextSearchMatch: me\.hasNonTextSearchMatch,/);
    assert.match(source, /options\.hasNonTextSearchMatch !== true &&/);
});

test('every search match must be checked, the select-all slot is not trusted (#99)', () => {
    const matches = [undefined, undefined, true, false, true];
    assert.equal(areAllSearchMatchesChecked(matches, [true, false, true, false, true]), true);
    assert.equal(areAllSearchMatchesChecked(matches, [true, false, true, true, false]), false);
    assert.equal(areAllSearchMatchesChecked([undefined, undefined, false], [true, false, false]), false);
    assert.match(source, /SSE\.Views\.AutoFilterDialog\.areAllSearchMatchesChecked\(me\.filterMatches, me\.filteredIndexes\)/);
});

test('a search with no listed match on a capped list can still be applied (#99b)', () => {
    assert.match(source, /me\.searchBeyondList = !!me\.filter && isCappedValueList && arr\.length === 0;/);
    assert.match(source, /if \(!me\.filter \|\| arr\.length>0 \|\| me\.searchBeyondList\)/);
    assert.match(source, /else if \(this\.searchBeyondList && this\.isSelectAllResultsChecked\(\)\)\s*\n\s*isValid = true;/);
    assert.match(source, /allMatchesChecked: me\.searchBeyondList \? me\.isSelectAllResultsChecked\(\)/);
});

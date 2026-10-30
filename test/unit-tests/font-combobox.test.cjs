// SPDX-FileCopyrightText: 2026 ElixirOffice contributors
// SPDX-License-Identifier: AGPL-3.0-only
'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const root = path.resolve(__dirname, '../..');

function loadCombo() {
    let specification;
    const context = {
        Asc: {FONT_THUMBNAIL_HEIGHT: 28},
        Common: {
            Controllers: {Desktop: {isActive: () => false}},
            Locale: {getCurrentLanguage: () => 'en'},
            Utils: {applicationPixelRatio: () => 1},
            UI: {
                ComboBox: {
                    extend(value) {
                        specification = value;
                        return function ComboBoxFonts() {};
                    }
                }
            }
        },
        document: {
            createElement() {
                return {getContext: () => ({})};
            }
        }
    };
    context.window = context;
    context._ = {extend: Object.assign, template: () => () => ''};
    context.$ = target => ({val: () => target.value});
    context.define = (_dependencies, factory) => factory();
    context.define.amd = true;
    vm.runInNewContext(
        fs.readFileSync(path.join(root, 'apps/common/main/lib/component/ComboBoxFonts.js'), 'utf8'),
        context,
        {filename: 'ComboBoxFonts.js'}
    );
    return specification;
}

function model(name) {
    return {
        get(key) {
            return key === 'name' ? name : undefined;
        },
        toJSON() {
            return {id: 'font-fira', name};
        }
    };
}

function select(value, available) {
    const events = [];
    const existing = model(available);
    const combo = {
        _skipInputChange: false,
        _isMouseDownMenu: false,
        _selectedItem: null,
        lastValue: 'Arial',
        valueField: 'name',
        displayField: 'name',
        store: {
            findWhere(query) {
                return query.name === available ? existing : undefined;
            }
        },
        setRawValue() {},
        trigger(name, _combo, record) {
            if (name === 'selected') events.push(record);
        },
        addItemToRecent() {},
        closeMenu() {}
    };
    loadCombo().onInputChanged.call(combo, {target: {value}, isDefaultPrevented: () => false});
    return events[0];
}

test('typed exact installed font resolves to its picker record', () => {
    const selected = select('Fira Sans', 'Fira Sans');
    assert.equal(selected.name, 'Fira Sans');
    assert.equal(selected.isNewFont, undefined);
});

test('unknown typed font remains explicitly new', () => {
    const selected = select('Not Installed', 'Fira Sans');
    assert.equal(selected.name, 'Not Installed');
    assert.equal(selected.isNewFont, true);
});

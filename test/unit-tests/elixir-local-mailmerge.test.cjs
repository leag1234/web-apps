// SPDX-FileCopyrightText: 2026 ElixirOffice contributors
// SPDX-License-Identifier: AGPL-3.0-only
'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const root = path.resolve(__dirname, '../..');

function loadController() {
    const Operation = require(path.join(root,
        'apps/common/main/lib/elixir-local/Operation.js'));
    const Images = {assertAvailable() {}, install() { return {dispose() {}}; }};
    const Password = {install() { return {
        dispose() {}, confirmExport() {}, isProtected() { return false; }
    }; }};
    const starts = [], ends = [], errors = [];
    let controller;
    const context = {
        console, TextDecoder, Uint8Array, URL, Map, Promise,
        crypto: require('node:crypto').webcrypto,
        Asc: {
            c_oAscFileType: {JSON: 2056},
            c_oAscAsyncAction: {MailMergeLoadFile: 20},
            c_oAscAsyncActionType: {BlockInteraction: 1}
        },
        AscCommon: {g_oDocumentUrls: {addImageUrl() {}}},
        EOWasmConvert: {formats: {docx: 65}},
        define(_dependencies, factory) {
            controller = factory(Images, {}, Password, Operation, {resolve() {}}, {});
        }
    };
    context.define.amd = true;
    vm.runInNewContext(fs.readFileSync(path.join(root,
        'apps/common/main/lib/elixir-local/Controller.js'), 'utf8'), context);
    const api = {
        DocInfo: {get_OfflineApp: () => true},
        asc_registerCallback() {}, asc_unregisterCallback() {},
        sync_StartAction(...values) { starts.push(values); },
        sync_EndAction(...values) { ends.push(values); }
    };
    const options = {format: 'docx', onExport() {}, onError: error => errors.push(error.message)};
    return {api, options, starts, ends, errors, local: controller.install(api, options)};
}

function load() {
    let mailMerge;
    const context = {
        TextDecoder,
        define(_dependencies, factory) {
            mailMerge = factory();
        }
    };
    vm.runInNewContext(fs.readFileSync(path.join(root,
        'apps/common/main/lib/elixir-local/MailMerge.js'), 'utf8'), context);
    return mailMerge;
}

test('recipient conversion uses x2t mail-merge JSON and returns at most 100 records', async () => {
    const calls = [];
    const converter = {convert: async (bytes, from, to, options) => {
        calls.push({bytes: Array.from(bytes), from, to, signal: options.signal});
        return new TextEncoder().encode(JSON.stringify([
            ['Name', 'City'], ...Array.from({length: 101}, (_, index) => ['Person ' + index, 'Paris'])
        ]));
    }};
    const signal = {aborted: false};
    const rows = await load().recipients(new Uint8Array([1, 2, 3]), 'xlsx', converter, {signal});
    assert.deepEqual(calls, [{bytes: [1, 2, 3], from: 'xlsx', to: 'json', signal}]);
    assert.equal(rows.length, 101);
    assert.deepEqual(Array.from(rows[0]), ['Name', 'City']);
    assert.deepEqual(Array.from(rows[100]), ['Person 99', 'Paris']);
});

test('recipient validation rejects unsupported, oversized, malformed and headerless inputs', async () => {
    const mailMerge = load();
    const converter = value => ({convert: async () => new TextEncoder().encode(value)});
    await assert.rejects(mailMerge.recipients(new Uint8Array([1]), 'ods', converter('[]')), /CSV or Excel/);
    await assert.rejects(mailMerge.recipients(new Uint8Array(20 * 1024 * 1024 + 1), 'csv', converter('[]')), /20 MiB/);
    await assert.rejects(mailMerge.recipients(new Uint8Array([1]), 'csv', converter('{')), /recipient data/);
    await assert.rejects(mailMerge.recipients(new Uint8Array([1]), 'csv', converter('[[]]')), /header/);
    await assert.rejects(mailMerge.recipients(new Uint8Array([1]), 'csv', converter('[["Name"],[{}]]')), /cell/);
});

test('merged output starts only after pending pictures settle', async () => {
    const mailMerge = load();
    let release;
    const pending = new Promise(resolve => { release = resolve; });
    const calls = [];
    assert.equal(mailMerge.afterPending(pending, () => calls.push('download'), error => { throw error; }, () => true), true);
    await Promise.resolve();
    assert.deepEqual(calls, []);
    release();
    await pending;
    await Promise.resolve();
    assert.deepEqual(calls, ['download']);
});

test('recipient import always ends its blocking action when another operation is active', async () => {
    const fixture = loadController();
    const active = fixture.local.beginOperation('other');
    fixture.api.mailMergeFileData = {fileType: 'csv', data: new Uint8Array([65])};
    fixture.api.downloadAs(20, {fileType: 2056});
    await new Promise(setImmediate);
    assert.equal(fixture.starts.length, 1);
    assert.equal(fixture.ends.length, 1);
    assert.match(fixture.errors[0], /already running/);
    assert(fixture.api.mailMergeFileData, 'a rejected import must retain its selected input');
    active.cancel();
    fixture.local.dispose();
});

test('recipient import always ends its blocking action when a save is pending', async () => {
    const fixture = loadController();
    let release;
    fixture.options.beforeOperationWork = () => new Promise(resolve => { release = resolve; });
    const saving = fixture.local.save().catch(() => {});
    fixture.api.mailMergeFileData = {fileType: 'csv', data: new Uint8Array([65])};
    fixture.api.downloadAs(20, {fileType: 2056});
    await new Promise(setImmediate);
    assert.equal(fixture.starts.length, 1);
    assert.equal(fixture.ends.length, 1);
    assert.match(fixture.errors[0], /already running/);
    fixture.local.dispose();
    release();
    await saving;
});

test('local Mailings exposes only the local recipient source', () => {
    const source = fs.readFileSync(path.join(root,
        'apps/documenteditor/main/app/controller/WordMailings.js'), 'utf8');
    assert.match(source, /txtFromUrl, value: 'url', visible: !mode\.isElixirLocal/);
    assert.match(source, /txtFromStorage, value: 'storage', visible: !!canStorage && !mode\.isElixirLocal/);
    assert.match(source, /isElixirLocal && this\.api\.eoLastMailMergeExport !== 'saved'/);
});

test('local controller reports merged output outcome and restores the native chooser hook', () => {
    const source = fs.readFileSync(path.join(root,
        'apps/common/main/lib/elixir-local/Controller.js'), 'utf8');
    assert.match(source, /chooseMailMergeRecipients/);
    assert.match(source, /eoLastMailMergeExport=acknowledgement&&acknowledgement\.kind\|\|'saved'/);
    assert.match(source, /eoLastMailMergeExport='pending'/);
    assert.match(source, /api\.asc_StartMailMerge=startMailMerge/);
});

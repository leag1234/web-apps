// SPDX-FileCopyrightText: 2026 ElixirOffice contributors
// SPDX-License-Identifier: AGPL-3.0-only
'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const root = path.resolve(__dirname, '../..');
const AUTOSTART_GUID = 'asc.{11111111-2222-4333-8444-555555555555}';

// The plugins controller, with one plugin to start automatically, and the
// launch state of the editor scripts, which the tests set.
function loadController() {
    const launch = { scriptsLoaded: false };
    const context = {
        // The reset logs an error when the toolbar has no plugins panel, which
        // these tests do not build.
        console: { error() {} },
        define(deps, factory) { factory(); },
        Backbone: { Controller: { extend: o => o } },
        _: { extend: (a, b) => Object.assign(a, b) },
        Common: { Controllers: { LaunchController: { isScriptLoaded: () => launch.scriptsLoaded } } }
    };
    context.define.amd = true;
    vm.runInNewContext(fs.readFileSync(path.join(root,
        'apps/common/main/lib/controller/Plugins.js'), 'utf8'), context);
    const started = [];
    const controller = Object.assign(Object.create(context.Common.Controllers.Plugins), {
        customButtonsArr: [],
        appOptions: {},
        autostart: [AUTOSTART_GUID],
        pluginsWinToShow: [],
        api: { asc_pluginRun: guid => started.push(guid) },
        // Registering the plugins with the SDK does not start them.
        refreshPluginsList() {}
    });
    return { controller, launch, started };
}

// A plugin store that holds plugins, as the editor resets it.
const store = { isEmpty: () => false, each() {} };

test('an autostart plugin waits for the editor scripts when the store is reset before them', () => {
    const { controller, launch, started } = loadController();

    controller.onResetPlugins(store);
    assert.deepEqual(started, []);

    launch.scriptsLoaded = true;
    controller.onPostLoadComplete();
    assert.deepEqual(started, [AUTOSTART_GUID]);
});

test('an autostart plugin starts at once when the store is reset after the editor scripts', () => {
    const { controller, launch, started } = loadController();
    launch.scriptsLoaded = true;

    controller.onResetPlugins(store);

    assert.deepEqual(started, [AUTOSTART_GUID]);
});

/*
 * SPDX-FileCopyrightText: 2026 ElixirOffice contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 *
 * ElixirShared: host context of the window chrome modules shared by the
 * editors (apps/common/main/lib/elixir-shared/). The editor-specific chrome
 * controller calls ElixirShared.init() before it creates the shared views:
 *
 *   ElixirShared.init(PE, 'pe', {
 *       formats: function () { return [...]; },   // Save As / Export cards
 *       defaultFormat: Asc.c_oAscFileType.PPTX,     // Share > Send a copy
 *       blankKey: 'txtBlankPresentation',           // first "New" thumbnail
 *       suggestions: function (toolbar) { return [btn, ...]; }, // Alt+Q, empty box
 *       protectController: null                     // Info > Restrict editing
 *   });
 *
 * Strings of the shared views are translated in each editor's locale files
 * under Common.Views.ElixirTitleBar.* and Common.Views.ElixirBackstage.*.
 */
define([], function () {
    'use strict';

    var Shared = window.ElixirShared = window.ElixirShared || {};

    Shared.init = function (app, prefix, opts) {
        Shared.app = app;
        Shared.prefix = prefix;
        Shared.opts = opts || {};
        return Shared;
    };
    Shared.opt = function (key, def) {
        var o = Shared.opts || {};
        return o[key] === undefined ? def : o[key];
    };

    return Shared;
});

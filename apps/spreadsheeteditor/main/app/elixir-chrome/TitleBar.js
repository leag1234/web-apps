/*
 * SPDX-FileCopyrightText: 2026 ElixirOffice contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 *
 * ElixirOffice window chrome: Excel-like title bar (spreadsheet editor).
 * The title bar is shared (common/main/lib/elixir-chrome/TitleBar.js); this
 * subclass carries the translations (locale SSE.Views.ElixirTitleBar.*) and
 * the commands suggested by the empty "Search (Alt+Q)" box, like Excel.
 */
define([
    'core',
    'common/main/lib/elixir-chrome/TitleBar'
], function () {
    'use strict';

    SSE.Views.ElixirTitleBar = Common.Views.ElixirTitleBar.extend(_.extend({
        settingsPrefix: 'sse',

        suggestionButtons: function () {
            var tb = this.toolbar || {};
            return [tb.btnInsertTable, tb.btnInsertChart, tb.btnCondFormat, tb.btnInsertFormula, tb.btnSortDown, tb.btnSetAutofilter, tb.btnFreezePanes];
        }
    }, SSE.Views.ElixirTitleBar || {}));
});

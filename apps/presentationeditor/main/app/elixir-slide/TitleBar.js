/*
 * SPDX-FileCopyrightText: 2026 ElixirOffice contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 *
 * ElixirOffice window chrome: PowerPoint-like title bar (presentation editor).
 * The title bar is shared (common/main/lib/elixir-chrome/TitleBar.js); this
 * subclass carries the translations (locale PE.Views.ElixirTitleBar.*) and
 * the commands suggested by the empty "Search (Alt+Q)" box.
 */
define([
    'core',
    'common/main/lib/elixir-chrome/TitleBar'
], function () {
    'use strict';

    PE.Views.ElixirTitleBar = Common.Views.ElixirTitleBar.extend(_.extend({
        settingsPrefix: 'pe',

        suggestionButtons: function () {
            var tb = this.toolbar || {};
            return [tb.btnAddSlide, tb.btnInsertTable, tb.btnsInsertImage && tb.btnsInsertImage[0],
                    tb.btnsInsertText && tb.btnsInsertText[0], tb.btnInsertChart, tb.btnSlideSize];
        }
    }, PE.Views.ElixirTitleBar || {}));
});

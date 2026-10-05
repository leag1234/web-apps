/*
 * ElixirOffice window chrome: Word-like title bar (document editor).
 *
 * The title bar itself is shared by the editors
 * (common/main/lib/elixir-chrome/TitleBar.js); this subclass carries the
 * document editor's translations (locale DE.Views.ElixirTitleBar.*) and the
 * commands suggested by the empty search box.
 *
 * Licensed under the GNU AGPL v3 like the rest of web-apps.
 */
define([
    'core',
    'common/main/lib/elixir-chrome/TitleBar'
], function () {
    'use strict';

    DE.Views.ElixirTitleBar = Common.Views.ElixirTitleBar.extend(_.extend({
        settingsPrefix: 'de',

        suggestionButtons: function () {
            var tb = this.toolbar || {};
            return [tb.btnInsertTable, tb.btnsPageBreak && tb.btnsPageBreak[0], tb.btnInsertShape, tb.btnInsertText, tb.btnPageMargins, tb.btnColumns];
        }
    }, DE.Views.ElixirTitleBar || {}));
});

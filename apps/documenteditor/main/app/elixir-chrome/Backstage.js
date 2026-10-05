/*
 * ElixirOffice window chrome: Word-like backstage (document editor).
 *
 * The backstage is shared by the editors
 * (common/main/lib/elixir-chrome/Backstage.js, Word formats by default); this
 * subclass carries the document editor's translations
 * (locale DE.Views.ElixirBackstage.*).
 *
 * Licensed under the GNU AGPL v3 like the rest of web-apps.
 */
define([
    'core',
    'common/main/lib/elixir-chrome/Backstage'
], function () {
    'use strict';

    DE.Views.ElixirBackstage = Common.Views.ElixirBackstage.extend(_.extend({}, DE.Views.ElixirBackstage || {}));
});

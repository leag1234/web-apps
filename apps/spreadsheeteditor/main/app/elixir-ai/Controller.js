/*
 * SPDX-FileCopyrightText: 2026 ElixirOffice contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 */
define(['common/main/lib/elixir-ai/Controller','spreadsheeteditor/main/app/elixir-ai/CellTools','spreadsheeteditor/main/app/elixir-ai/CellStrings'], function (_controller, _tools, English) {
    'use strict';
    // Each editor has its own window. Retain the established controller lookup used by
    // native Options, the shared title bar and File backstage.
    var AI=Common.ElixirAI;
    AI.CellStrings=AI.CellStrings||{};
    AI.cellText=function(key,values){var value=AI.CellStrings[key]||English[key]||key;return value.replace(/\{(\w+)\}/g,function(m,k){return values&&values[k]!==undefined?String(values[k]):m;});};
    Common.Controllers.ElixirAI=Common.Controllers.ElixirAI.extend({toolsetId:'cell'});
    return Common.Controllers.ElixirAI;
});

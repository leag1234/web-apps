// SPDX-License-Identifier: AGPL-3.0-only
define([], function () {
    'use strict';
    return Object.freeze({
        placeholder:'Ask Elixir… e.g. “Write a total formula”, “Sort sales by amount”',
        footer:'Context sent: the selected range, or the used table when one cell is selected. Reads are bounded.',
        scope:'Selected range or used table',
        scopeRange:'{sheet} — {range}',
        read:'Reading cells and formulas',
        explain:'Reading the formula to explain it',
        summarize:'Summarizing the range',
        formula:'Wrote formulas in {range}',
        trim:'Trimmed whitespace in {range}',
        split:'Split text into {range}',
        dedupe:'Removed duplicate rows inside {range}',
        dates:'Normalized dates in {range}',
        numbers:'Normalized numbers in {range}',
        sort:'Sorted complete rows in {range}',
        filter:'Filtered rows in {range}',
        chart:'Created a chart from {range}',
        pivot:'Created a pivot table from {range}',
        rollbackUnverified:'The action failed. Cancellation could not be verified. Review your spreadsheet and use Undo if needed.',
        staleDiscard:'The spreadsheet changed after this action. Use Undo to review later changes first; no changes were discarded.'
    });
});

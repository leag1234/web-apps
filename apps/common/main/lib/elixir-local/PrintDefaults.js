/*
 * SPDX-FileCopyrightText: 2026 ElixirOffice contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 */
// Smart defaults of the spreadsheet print panel (pure logic, no editor objects, lengths in mm).
// They only shape the print run and the panel state: the document keeps its stored page setup until the user saves it.
(function (factory) {
    if (typeof define === 'function' && define.amd) define([], factory);
    else if (typeof module === 'object' && module.exports) module.exports = factory();
    else (globalThis.Common = globalThis.Common || {}).ElixirPrintDefaults = factory();
})(function () {
    'use strict';
    var PORTRAIT = 0, LANDSCAPE = 1;
    var DEFAULT_MARGINS = {left: 17.8, right: 17.8, top: 19.1, bottom: 19.1};

    function num(value) { return typeof value === 'number' && isFinite(value) ? value : null; }

    // The sheet is "untouched" when it prints at the default 100% without any fit-to setting.
    function hasDefaultScale(stored) {
        var scale = num(stored.scale);
        return (scale === null || scale === 100) && !stored.fitToWidth && !stored.fitToHeight;
    }

    function printableWidth(pageWidth, pageHeight, orientation, margins) {
        var width = orientation === LANDSCAPE ? pageHeight : pageWidth;
        return width - margins.left - margins.right;
    }

    /**
     * @param {Object} input
     * @param {{scale, fitToWidth, fitToHeight, orientation, horizontalCentered}} input.stored page setup stored for the sheet
     * @param {number} input.pageWidth page width in mm for portrait (as stored); input.pageHeight likewise
     * @param {{left,right,top,bottom}} [input.margins] margins actually used for printing (mm)
     * @param {number} input.rangeWidth width of the print range in mm; input.rangeHeight its height (null when unknown)
     * @param {boolean} input.hasPrintArea the sheet has a user defined print area (and it is not ignored)
     * @returns {{fitToWidth: boolean, orientation: (number|undefined), horizontalCentered: boolean, reasons: string[]}} what to preselect
     */
    // A stored scale this small with no fit-to setting is a symptom (e.g. a fit computed on broken margins), not a choice.
    var TINY_SCALE = 25;
    function hasTinyScale(stored) {
        var scale = num(stored.scale);
        return scale !== null && scale < TINY_SCALE && !stored.fitToWidth && !stored.fitToHeight;
    }

    // One page per sheet is the default when the text stays readable (scale >= 45%); otherwise fit to the page width only.
    var MIN_ONE_PAGE_SCALE = 45;

    function printableHeight(pageWidth, pageHeight, orientation, margins) {
        var height = orientation === LANDSCAPE ? pageWidth : pageHeight;
        return height - margins.top - margins.bottom;
    }

    // scale (0..) that fits the whole range on one page in the given orientation
    function onePageScale(o, pageW, pageH, margins, rangeW, rangeH) {
        var w = printableWidth(pageW, pageH, o, margins), h = printableHeight(pageW, pageH, o, margins);
        if (!(w > 0) || !(h > 0)) return 0;
        return Math.min(w / rangeW, rangeH > 0 ? h / rangeH : Infinity);
    }

    function decide(input) {
        var result = {fitToWidth: false, fitToHeight: false, orientation: undefined, horizontalCentered: false, tinyScale: false, scale: null, reasons: []};
        var stored = input.stored || {}, rangeW = num(input.rangeWidth), rangeH = num(input.rangeHeight),
            pageW = num(input.pageWidth), pageH = num(input.pageHeight);
        var tiny = hasTinyScale(stored);
        var sizesKnown = rangeW !== null && rangeH !== null && pageW !== null && pageH !== null && rangeW > 0;
        if (!tiny) {
            if (input.hasPrintArea || !sizesKnown) return result;          // a user range is left exactly as set
            if (!hasDefaultScale(stored)) return result;                   // the user chose a scale or fit-to: keep it
        }
        var m = input.margins || DEFAULT_MARGINS;
        var margins = {left: num(m.left) === null ? DEFAULT_MARGINS.left : m.left, right: num(m.right) === null ? DEFAULT_MARGINS.right : m.right,
            top: num(m.top) === null ? DEFAULT_MARGINS.top : m.top, bottom: num(m.bottom) === null ? DEFAULT_MARGINS.bottom : m.bottom};
        var orientation = stored.orientation === LANDSCAPE ? LANDSCAPE : PORTRAIT;
        var known = sizesKnown && !input.hasPrintArea;
        var onePage = false;
        if (known) {
            var current = onePageScale(orientation, pageW, pageH, margins, rangeW, rangeH);
            if (!tiny && current >= 1) return result;                      // already one page at 100%
            var best = current;
            var sideways0 = orientation === PORTRAIT ? onePageScale(LANDSCAPE, pageW, pageH, margins, rangeW, rangeH) : 0;
            if (!tiny && current === 0 && sideways0 === 0) return result;   // margins leave no printable area: nothing to guess
            if (orientation === PORTRAIT) {                                // portrait is the default; sideways only if it really helps
                var sideways = onePageScale(LANDSCAPE, pageW, pageH, margins, rangeW, rangeH);
                if (sideways > current) { best = sideways; result.orientation = LANDSCAPE; }
            }
            if (best * 100 >= MIN_ONE_PAGE_SCALE) {
                onePage = true;
                result.scale = Math.min(best, 1);
            } else {
                // too small for one page: fit to the page width across several pages (sideways when wider than tall and not fitting portrait)
                result.orientation = undefined;
                var width = printableWidth(pageW, pageH, orientation, margins);
                if (orientation === PORTRAIT && rangeW > rangeH && width > 0 && rangeW > width) result.orientation = LANDSCAPE;
            }
            if (result.orientation === LANDSCAPE) result.reasons.push('landscape');
        }
        result.fitToWidth = true;
        result.fitToHeight = onePage;
        result.tinyScale = tiny;
        result.reasons.push(onePage ? 'one-page' : tiny && !known ? 'tiny-scale' : tiny ? 'tiny-scale' : 'fit-width');
        if (!stored.horizontalCentered) {
            result.horizontalCentered = true;
            result.reasons.push('center');
        }
        return result;
    }

    // Width/height of a range from per-column widths and per-row heights (mm), for engines that report no size.
    function sum(list) { var total = 0; for (var i = 0; i < list.length; i++) total += list[i] > 0 ? list[i] : 0; return total; }

    // Printed messages of the engine warnings, filled in with the sheet name and values.
    function formatWarning(warning, templates) {
        var text = warning.kind === 'margins' ? templates.margins : warning.kind === 'fitted' ? templates.fitted : templates.scale;
        return text.replace('{name}', warning.name).replace('{scale}', String(Math.round(warning.scale)));
    }

    // Normalises the result of asc_GetPrintSetupWarnings or the older asc_GetPrintScaleWarnings.
    function normalizeWarnings(raw) {
        if (!Array.isArray(raw)) return [];
        return raw.filter(function (w) { return w && typeof w === 'object'; }).map(function (w) {
            return {kind: w.kind === 'margins' ? 'margins' : 'scale', sheetIndex: w.sheetIndex, name: String(w.name === undefined ? '' : w.name), scale: w.scale};
        });
    }

    return {decide: decide, hasDefaultScale: hasDefaultScale, hasTinyScale: hasTinyScale, TINY_SCALE: TINY_SCALE, MIN_ONE_PAGE_SCALE: MIN_ONE_PAGE_SCALE, sum: sum, formatWarning: formatWarning, normalizeWarnings: normalizeWarnings,
        PORTRAIT: PORTRAIT, LANDSCAPE: LANDSCAPE, DEFAULT_MARGINS: DEFAULT_MARGINS};
});

/*
 * SPDX-FileCopyrightText: 2026 ElixirOffice contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 */
// One print panel, two entry points: File > Print opens it as "Print", File > Export as PDF (and the PDF choice of
// Save As / Export a copy) opens it as "Export as PDF". The desktop shell sets the mode on the editor's
// InternalSettings ('eo-print-mode') right before it triggers 'file:print'; the panel reads and clears it when shown.
(function (factory) {
    if (typeof define === 'function' && define.amd) define([], factory);
    else if (typeof module === 'object' && module.exports) module.exports = factory();
    else (globalThis.Common = globalThis.Common || {}).ElixirPrintMode = factory();
})(function () {
    'use strict';
    var KEY = 'eo-print-mode';

    // returns 'pdf' or 'print' and resets the request so a later Ctrl+P is a plain Print
    function consume(settings) {
        var mode = 'print';
        try {
            if (settings && settings.get(KEY) === 'pdf') mode = 'pdf';
            settings && settings.set(KEY, 'print');
        } catch (e) { /* settings unavailable: plain print */ }
        return mode;
    }

    function toggle(button, on) {
        var el = button && (button.cmpEl || button.$el);
        if (el && el.toggleClass) el.toggleClass('primary', !!on);
    }

    // view: the print panel (single buttons btnPrint/btnPrintPdf, or arrays btnsPrint/btnsPrintPDF)
    function apply(view, mode, strings) {
        var pdf = mode === 'pdf';
        var title = view.$el && view.$el.find && view.$el.find('.main-header');
        if (title && title.length) title.text(pdf ? strings.exportPdf : strings.print);
        [].concat(view.btnPrint || view.btnsPrint || []).forEach(function (b) { toggle(b, !pdf); });
        [].concat(view.btnPrintPdf || view.btnsPrintPDF || []).forEach(function (b) { toggle(b, pdf); });
        return pdf;
    }

    return {KEY: KEY, consume: consume, apply: apply};
});

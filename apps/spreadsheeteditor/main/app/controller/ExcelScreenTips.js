/*
 * SPDX-FileCopyrightText: 2026 ElixirOffice contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 *
 * SSE.Controllers.ExcelScreenTips: Excel's ScreenTips (command name with its
 * shortcut, then one line saying what it does) on the spreadsheet ribbon; the
 * engine is shared with the document editor
 * (common/main/lib/controller/ElixirScreenTips.js).
 */

define([
    'core',
    'common/main/lib/controller/ElixirScreenTips'
], function () {
    'use strict';

    var S = function (id) { return '#' + id; };

    var TIPS = [
        [S('slot-btn-paste'), 'descPaste'], [S('slot-btn-cut'), 'descCut'], [S('slot-btn-copy'), 'descCopy'], [S('slot-btn-copystyle'), 'descFormatPainter'],
        [S('slot-field-fontname'), 'descFont'], [S('slot-field-fontsize'), 'descFontSize'], [S('slot-btn-incfont'), 'descGrowFont'], [S('slot-btn-decfont'), 'descShrinkFont'],
        [S('slot-btn-bold'), 'descBold'], [S('slot-btn-italic'), 'descItalic'], [S('slot-btn-underline'), 'descUnderline'],
        [S('slot-btn-borders'), 'descBorders'], [S('slot-btn-fillparag'), 'descFillColor'], [S('slot-btn-fontcolor'), 'descFontColor'],
        [S('slot-btn-top'), 'descTopAlign'], [S('slot-btn-middle'), 'descMiddleAlign'], [S('slot-btn-bottom'), 'descBottomAlign'],
        [S('slot-btn-text-orient'), 'descOrientation'], [S('slot-btn-wrap'), 'descWrapText'],
        [S('slot-btn-align-left'), 'descAlignLeft'], [S('slot-btn-align-center'), 'descCenter'], [S('slot-btn-align-right'), 'descAlignRight'],
        [S('slot-elx-indent-dec'), 'descDecreaseIndent'], [S('slot-elx-indent-inc'), 'descIncreaseIndent'], [S('slot-btn-merge'), 'descMerge'],
        [S('slot-btn-format'), 'descNumberFormat'], [S('slot-btn-currency'), 'descAccounting'], [S('slot-btn-percents'), 'descPercent'],
        [S('slot-btn-comma'), 'descComma'], [S('slot-btn-digit-inc'), 'descIncDecimal'], [S('slot-btn-digit-dec'), 'descDecDecimal'],
        [S('slot-btn-condformat'), 'descCondFormat'], [S('slot-btn-table-tpl'), 'descFormatAsTable'],
        [S('slot-btn-cell-ins'), 'descInsertCells'], [S('slot-btn-cell-del'), 'descDeleteCells'], [S('slot-btn-cell-format'), 'descFormatCells'],
        [S('slot-btn-formula'), 'descAutoSum'], [S('slot-btn-fill-num'), 'descFill'], [S('slot-btn-clear'), 'descClear'],
        [S('slot-elx-sort-filter'), 'descSortFilter'], [S('slot-elx-find-select'), 'descFindSelect'],
        [S('slot-btn-instable'), 'descTable'], ['.panel[data-tab=ins] .slot-add-pivot', 'descPivot'], [S('slot-btn-insimage'), 'descPicture'],
        [S('slot-btn-insshape'), 'descShapes'], [S('slot-btn-insrecommend'), 'descRecommendedCharts'], [S('slot-btn-inschart'), 'descChart'],
        [S('slot-btn-inssparkline'), 'descSparkline'], [S('slot-btn-insslicer'), 'descSlicer'], [S('slot-btn-inshyperlink'), 'descLink'],
        [S('slot-btn-instext'), 'descTextBox'], [S('slot-btn-insequation'), 'descEquation'], [S('slot-btn-inssymbol'), 'descSymbol'],
        [S('slot-btn-inscheckbox'), 'descCheckbox'],
        [S('slot-btn-pagemargins'), 'descMargins'], [S('slot-btn-pageorient'), 'descOrientationPage'], [S('slot-btn-pagesize'), 'descSize'],
        [S('slot-btn-printarea'), 'descPrintArea'], [S('slot-btn-pagebreak'), 'descBreaks'], [S('slot-btn-printtitles'), 'descPrintTitles'],
        [S('slot-btn-additional-formula'), 'descInsertFunction'], [S('slot-btn-autosum'), 'descAutoSum'], [S('slot-btn-named-range-huge'), 'descNames'],
        [S('slot-btn-trace-prec'), 'descTracePrecedents'], [S('slot-btn-trace-dep'), 'descTraceDependents'], [S('slot-btn-show-formulas'), 'descShowFormulas'],
        [S('slot-btn-watch-window'), 'descWatchWindow'], [S('slot-btn-calculate'), 'descCalculation'],
        [S('slot-btn-custom-sort'), 'descSort'], [S('slot-btn-text-column'), 'descTextToColumns'], [S('slot-btn-rem-duplicates'), 'descRemoveDuplicates'],
        [S('slot-btn-data-validation'), 'descDataValidation'], [S('slot-btn-goal-seek'), 'descGoalSeek'], [S('slot-btn-group'), 'descGroup'],
        [S('slot-btn-ungroup'), 'descUngroup'], [S('slot-btn-freeze'), 'descFreezePanes'],
        ['header.btnUndo', 'descUndo'], ['header.btnRedo', 'descRedo'], ['header.btnSave', 'descSave'], ['header.btnPrint', 'descPrint']
    ];

    SSE.Controllers.ExcelScreenTips = Common.Controllers.ElixirScreenTips.extend(_.extend({
        tips: TIPS,

        // the tips go on the controls inside the slots
        resolve: function (path) {
            if (/^[#.]/.test(path)) return $('#toolbar').find(path).first().children().first();
            return Common.Controllers.ElixirScreenTips.prototype.resolve.call(this, path);
        },

        descPaste: 'Add content on the Clipboard to your workbook.',
        descCut: 'Remove the selection and put it on the Clipboard so you can paste it somewhere else.',
        descCopy: 'Put a copy of the selection on the Clipboard so you can paste it somewhere else.',
        descFormatPainter: 'Like the look of one selection? You can apply that look to other cells.',
        descFont: 'Pick a new font for your text.',
        descFontSize: 'Change the size of your text.',
        descGrowFont: 'Make your text a bit bigger.',
        descShrinkFont: 'Make your text a bit smaller.',
        descBold: 'Make your text bold.',
        descItalic: 'Italicize your text.',
        descUnderline: 'Underline your text.',
        descBorders: 'Apply borders to the selected cells.',
        descFillColor: 'Color the background of cells to make them stand out.',
        descFontColor: 'Change the color of your text.',
        descTopAlign: 'Align text to the top of the cell.',
        descMiddleAlign: 'Align text so that it is centered between the top and bottom of the cell.',
        descBottomAlign: 'Align text to the bottom of the cell.',
        descOrientation: 'Rotate your text diagonally or vertically.',
        descWrapText: 'Wrap extra-long text into multiple lines so you can see all of it.',
        descAlignLeft: 'Align your content to the left.',
        descCenter: 'Center your content.',
        descAlignRight: 'Align your content to the right.',
        descDecreaseIndent: 'Move your content closer to the cell border.',
        descIncreaseIndent: 'Move your content farther away from the cell border.',
        descMerge: 'Combine and center the contents of the selected cells in a new larger cell.',
        descNumberFormat: 'Choose the format for your cells, such as percentage, currency, date or time.',
        descAccounting: 'Choose an alternate currency format for the selected cell.',
        descPercent: 'Format as a percent.',
        descComma: 'Format with a thousands separator.',
        descIncDecimal: 'Show more decimal places for a more precise value.',
        descDecDecimal: 'Show fewer decimal places.',
        descCondFormat: 'Easily spot trends and patterns in your data using bars, colors and icons to visually highlight important values.',
        descFormatAsTable: 'Quickly convert a range of cells to a table with its own style.',
        descInsertCells: 'Add cells, rows, columns or sheets to your workbook.',
        descDeleteCells: 'Delete cells, rows, columns or sheets from your workbook.',
        descFormatCells: 'Change the row height or column width, organize sheets, or protect or hide cells.',
        descAutoSum: 'Automatically add a quick calculation to your worksheet, such as a sum or an average.',
        descFill: 'Continue a pattern into one or more adjacent cells.',
        descClear: 'Delete everything in the cell, or remove just the formatting, contents, comments or hyperlinks.',
        descSortFilter: 'Organize your data so it is easier to analyze.',
        descFindSelect: 'Find, replace and select cells in your workbook.',
        descTable: 'Create a table to organize and analyze related data.',
        descPivot: 'Easily arrange and summarize complex data in a PivotTable.',
        descPicture: 'Insert pictures from your computer or from the web.',
        descShapes: 'Insert ready-made shapes, such as circles, squares and arrows.',
        descRecommendedCharts: 'Want us to recommend a good chart to showcase your data? Just select data in your worksheet and click this button.',
        descChart: 'Add a chart to visualize and compare your data.',
        descSparkline: 'Insert a tiny chart in a single cell to show a trend.',
        descSlicer: 'Use a visual slicer to filter data quickly.',
        descLink: 'Create a link in your workbook for quick access to webpages and files.',
        descTextBox: 'Draw a text box anywhere.',
        descEquation: 'Add common mathematical equations to your workbook.',
        descSymbol: 'Add symbols that are not on your keyboard.',
        descCheckbox: 'Insert checkboxes in the selected cells.',
        descMargins: 'Set the margins for the entire document or the current section.',
        descOrientationPage: 'Give your pages a portrait or landscape layout.',
        descSize: 'Choose a paper size for your printout.',
        descPrintArea: 'Mark a specific area of the sheet for printing.',
        descBreaks: 'Add a break where you want the next page to begin in the printed copy.',
        descPrintTitles: 'Pick rows and columns to repeat on each printed page.',
        descInsertFunction: 'Work with the formula in the current cell. You can easily pick functions to use and get help on how to fill out the input values.',
        descNames: 'Create, edit, delete and find all the names used in the workbook.',
        descTracePrecedents: 'Show arrows that indicate which cells affect the value of the selected cell.',
        descTraceDependents: 'Show arrows that indicate which cells are affected by the value of the selected cell.',
        descShowFormulas: 'Display the formula in each cell instead of the resulting value.',
        descWatchWindow: 'Keep an eye on the values of some cells as you make changes to the sheet.',
        descCalculation: 'Choose to calculate formulas automatically or manually.',
        descSort: 'Find values quickly by sorting your data.',
        descTextToColumns: 'Take the text in one or more cells and split it into multiple cells.',
        descRemoveDuplicates: 'Delete duplicate rows from a sheet.',
        descDataValidation: 'Pick from a list of rules to limit the type of data that can be entered in a cell.',
        descGoalSeek: 'Find the input value that gives the result you want.',
        descGroup: 'Tie a range of cells together so they can be collapsed or expanded.',
        descUngroup: 'Ungroup a range of cells that were previously grouped.',
        descFreezePanes: 'Keep a portion of the sheet visible while the rest of the sheet scrolls.',
        descUndo: 'Undo your last action.',
        descRedo: 'Redo the action you undid.',
        descSave: 'Save your workbook.',
        descPrint: 'Print your workbook.'
    }, SSE.Controllers.ExcelScreenTips || {}));
});

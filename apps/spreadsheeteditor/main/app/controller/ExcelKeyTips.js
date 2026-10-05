/*
 * SPDX-FileCopyrightText: 2026 ElixirOffice contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 *
 * SSE.Controllers.ExcelKeyTips: Excel's KeyTip letters (Excel for Microsoft
 * 365, en-US) for the spreadsheet ribbon; the engine is shared with the
 * document editor (common/main/lib/controller/ElixirKeyTips.js).
 * Accessors are the ribbon slots of Toolbar.template (CSS selectors).
 */

define([
    'core',
    'common/main/lib/controller/ElixirKeyTips'
], function () {
    'use strict';

    var TABS = {
        file: 'F', home: 'H', ins: 'N', draw: 'JI', layout: 'P', formula: 'M', data: 'A',
        review: 'R', view: 'W', plugins: 'X', pivot: 'JT', tabledesign: 'JT',
        chartdesign: 'JC', chartformat: 'JO', shapeformat: 'JD', pictureformat: 'JP'
    };

    var S = function (id) { return '#' + id; },
        IN = function (tab, sel) { return '.panel[data-tab=' + tab + '] ' + sel; };

    var COMMANDS = {
        home: [
            [S('slot-btn-paste'), 'V'], [S('slot-btn-cut'), 'X'], [S('slot-btn-copy'), 'C'], [S('slot-btn-copystyle'), 'FP'],
            [S('slot-field-fontname'), 'FF'], [S('slot-field-fontsize'), 'FS'], [S('slot-btn-incfont'), 'FG'], [S('slot-btn-decfont'), 'FK'],
            [S('slot-btn-bold'), '1'], [S('slot-btn-italic'), '2'], [S('slot-btn-underline'), '3'], [S('slot-btn-borders'), 'B'],
            [S('slot-btn-fillparag'), 'H'], [S('slot-btn-fontcolor'), 'FC'],
            [S('slot-btn-top'), 'AT'], [S('slot-btn-middle'), 'AM'], [S('slot-btn-bottom'), 'AB'], [S('slot-btn-text-orient'), 'FQ'],
            [S('slot-btn-wrap'), 'W'], [S('slot-btn-align-left'), 'AL'], [S('slot-btn-align-center'), 'AC'], [S('slot-btn-align-right'), 'AR'],
            [S('slot-elx-indent-dec'), '5'], [S('slot-elx-indent-inc'), '6'], [S('slot-btn-merge'), 'M'],
            [S('slot-btn-format'), 'N'], [S('slot-btn-currency'), 'AN'], [S('slot-btn-percents'), 'P'], [S('slot-btn-comma'), 'K'],
            [S('slot-btn-digit-inc'), '0'], [S('slot-btn-digit-dec'), '9'],
            [S('slot-btn-condformat'), 'L'], [S('slot-btn-table-tpl'), 'T'], [S('slot-field-styles'), 'J'], ['.elx-collapsed-group', 'J'],
            [S('slot-btn-cell-ins'), 'I'], [S('slot-btn-cell-del'), 'D'], [S('slot-btn-cell-format'), 'O'],
            [S('slot-btn-formula'), 'U'], [S('slot-btn-fill-num'), 'FI'], [S('slot-btn-clear'), 'E'],
            [S('slot-elx-sort-filter'), 'S'], [S('slot-elx-find-select'), 'FD']
        ],
        ins: [
            [IN('ins', '.slot-add-pivot'), 'V'], [S('slot-btn-instable'), 'T'], [S('slot-btn-insimage'), 'P'], [S('slot-btn-insshape'), 'SH'],
            [S('slot-btn-inssmartart'), 'M1'], [S('slot-btn-insrecommend'), 'R'], [S('slot-btn-inschart'), 'C'], [S('slot-btn-inssparkline'), 'SL'],
            [S('slot-btn-insslicer'), 'SF'], [S('slot-btn-inshyperlink'), 'I'], [IN('ins', '.slot-comment'), 'M'],
            [S('slot-btn-instext'), 'X'], [IN('ins', '.slot-editheader'), 'H'], [S('slot-btn-instextart'), 'W'],
            [S('slot-btn-insequation'), 'E'], [S('slot-btn-inssymbol'), 'U'], [S('slot-btn-inscheckbox'), 'CB']
        ],
        layout: [
            [S('slot-btn-colorschemas'), 'TC'], [S('slot-btn-pagemargins'), 'M'], [S('slot-btn-pageorient'), 'O'], [S('slot-btn-pagesize'), 'SZ'],
            [S('slot-btn-printarea'), 'R'], [S('slot-btn-pagebreak'), 'B'], [IN('layout', '.slot-editheader'), 'HF'], [S('slot-btn-printtitles'), 'I'],
            [S('slot-btn-scale'), 'SC'], [S('slot-chk-print-gridlines'), 'PG'], [S('slot-chk-print-headings'), 'PH'],
            [S('slot-img-movefrwd'), 'AF'], [S('slot-img-movebkwd'), 'AE'], [S('slot-img-align'), 'AA'], [S('slot-img-group'), 'GP']
        ],
        formula: [
            [S('slot-btn-additional-formula'), 'F'], [S('slot-btn-autosum'), 'U'], [S('slot-btn-recent'), 'R'], [S('slot-btn-financial'), 'I'],
            [S('slot-btn-logical'), 'L'], [S('slot-btn-text'), 'T'], [S('slot-btn-datetime'), 'E'], [S('slot-btn-lookup'), 'O'],
            [S('slot-btn-math'), 'G'], [S('slot-btn-more'), 'Q'], [S('slot-btn-named-range-huge'), 'N'],
            [S('slot-btn-trace-prec'), 'P'], [S('slot-btn-trace-dep'), 'D'], [S('slot-btn-remove-arrows'), 'A'], [S('slot-btn-show-formulas'), 'H'],
            [S('slot-btn-watch-window'), 'W'], [S('slot-btn-calculate'), 'X']
        ],
        data: [
            [S('slot-btn-data-from-text'), 'PT'], [S('slot-btn-data-external-links'), 'K'],
            [IN('data', '.slot-sortdesc'), 'SA'], [IN('data', '.slot-sortasc'), 'SD'], [S('slot-btn-custom-sort'), 'SS'],
            [IN('data', '.slot-btn-setfilter'), 'T'], [IN('data', '.slot-btn-clear-filter'), 'C'],
            [S('slot-btn-text-column'), 'E'], [S('slot-btn-rem-duplicates'), 'M'], [S('slot-btn-data-validation'), 'V'],
            [S('slot-btn-goal-seek'), 'W'], [S('slot-btn-group'), 'G'], [S('slot-btn-ungroup'), 'U'], [S('slot-btn-solver'), 'Y2']
        ],
        review: [
            [IN('review', '.slot-comment'), 'C'], [S('slot-comment-remove'), 'D'],
            [S('slot-btn-protect-sheet'), 'PS'], [S('slot-btn-protect-wb'), 'PW'], [S('slot-btn-allow-ranges'), 'U']
        ],
        view: [
            [S('slot-btn-view-normal'), 'L'], [S('slot-btn-view-pagebreak'), 'I'], [S('slot-field-zoom'), 'Q'],
            [S('slot-btn-freeze'), 'F'], [S('slot-chk-gridlines'), 'VG'], [S('slot-chk-heading'), 'VH'], [S('slot-chk-formula'), 'VF'],
            [S('slot-btn-macros'), 'M']
        ]
    };

    // Excel's backstage letters
    var BACKSTAGE = {
        'fm-btn-elx-home': 'H', 'fm-btn-create': 'N', 'fm-btn-elx-open': 'O', 'fm-btn-info': 'I', 'fm-btn-save': 'S',
        'fm-btn-save-copy': 'Y', 'fm-btn-download': 'A', 'fm-btn-print-with-preview': 'P', 'fm-btn-elx-share': 'Z',
        'fm-btn-elx-export': 'E', 'fm-btn-elx-close': 'C', 'fm-btn-settings': 'T', 'fm-btn-help': 'H1', 'fm-btn-elx-about': 'D'
    };

    SSE.Controllers.ExcelKeyTips = Common.Controllers.ElixirKeyTips.extend({
        tabs: TABS,
        commands: COMMANDS,
        backstage: BACKSTAGE
    });
});

/*
 * SPDX-FileCopyrightText: 2026 ElixirOffice contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 *
 * PE.Controllers.PowerPointKeyTips: PowerPoint's KeyTips (Alt, the tab
 * letter, then the command letters) on the shared KeyTips controller
 * (common/main/lib/controller/ElixirKeyTips.js). Controls without a
 * PowerPoint equivalent keep the letter the hint manager assigns.
 */
define([
    'core',
    'common/main/lib/controller/ElixirKeyTips'
], function () {
    'use strict';

    var TABS = {
        file: 'F', home: 'H', ins: 'N', draw: 'JI', design: 'G', transit: 'K', animate: 'A', slideshow: 'S',
        review: 'R', view: 'W', slideMaster: 'M', plugins: 'X',
        shapeformat: 'JD', pictureformat: 'JP', tabledesign: 'JT', tablelayout: 'JL', chartdesign: 'JC'
    };

    var S = function (id) { return '#' + id; },
        IN = function (tab, sel) { return '.panel[data-tab=' + tab + '] ' + sel; };

    var COMMANDS = {
        home: [
            [S('slot-btn-paste'), 'V'], [S('slot-btn-cut'), 'X'], [S('slot-btn-copy'), 'C'], [S('slot-btn-copystyle'), 'FP'],
            [S('slot-addslide'), 'I'], [S('slot-changeslide'), 'L'], [S('slot-elx-reset-slide'), 'RE'],
            [S('slot-field-fontname'), 'FF'], [S('slot-field-fontsize'), 'FS'], [S('slot-btn-incfont'), 'FG'], [S('slot-btn-decfont'), 'FK'],
            [S('slot-btn-clearstyle'), 'E'], [S('slot-btn-bold'), '1'], [S('slot-btn-italic'), '2'], [S('slot-btn-underline'), '3'],
            [S('slot-btn-strikeout'), '4'], [S('slot-btn-changecase'), '7'], [S('slot-btn-fontcolor'), 'FC'],
            [S('slot-btn-markers'), 'U'], [S('slot-btn-numbering'), 'N'], [S('slot-btn-decoffset'), 'AO'], [S('slot-btn-incoffset'), 'AI'],
            [S('slot-btn-linespace'), 'K'], [S('slot-btn-columns'), 'AC'], [S('slot-btn-direction'), 'AD'], [S('slot-btn-halign'), 'AL'],
            [S('slot-btn-valign'), 'AT'], [S('slot-btn-arrange-shape'), 'G'], [S('slot-btn-align-shape'), 'AA'],
            [S('slot-elx-find'), 'FD'], [S('slot-btn-replace'), 'R'], [S('slot-btn-select-all'), 'SL']
        ],
        ins: [
            [S('slot-elx-ins-newslide'), 'I'], [S('slot-btn-inserttable'), 'T'], [IN('ins', '.slot-insertimg'), 'P'],
            [S('slot-btn-inssmartart'), 'M'], [S('slot-btn-insertchart'), 'C'], [S('slot-btn-insertlink'), 'IK'],
            [IN('ins', '.slot-comment'), 'L'], [IN('ins', '.slot-instext'), 'X'], [S('slot-btn-editheader'), 'H'],
            [S('slot-btn-instextart'), 'W'], [S('slot-btn-datetime'), 'D'], [S('slot-btn-slidenum'), 'SN'],
            [S('slot-btn-insertequation'), 'E'], [S('slot-btn-inssymbol'), 'U'], [S('slot-btn-insvideo'), 'V'], [S('slot-btn-insaudio'), 'O']
        ],
        design: [[S('slot-btn-colorschemas'), 'C'], [S('slot-btn-slidesize'), 'S'], [S('slot-elx-format-bg'), 'G']],
        transit: [[S('transit-button-preview'), 'P'], [S('transit-button-parameters'), 'EO'], [S('transit-button-apply'), 'L']],
        animate: [[S('animation-button-preview'), 'P'], [S('animation-button-parameters'), 'EM'], [S('animation-button-add-effect'), 'AA'],
                  [S('animation-button-pane'), 'C'], [S('animation-trigger'), 'AT'], [S('animation-moveearlier'), 'E'], [S('animation-movelater'), 'L']],
        slideshow: [[S('slot-elx-show-begin'), 'B'], [S('slot-elx-show-current'), 'C'], [S('slot-elx-show-presenter'), 'PR'],
                    [S('slot-elx-show-setup'), 'S'], [S('slot-elx-show-hide'), 'H'], [S('slot-elx-show-usepresenter'), 'V']],
        view: [[S('slot-btn-normal'), 'L'], [S('slot-btn-slide-master'), 'M'], [S('slot-chk-notes'), 'PN'], [S('slot-chk-rulers'), 'R'],
               [S('slot-btn-guides'), 'SG'], [S('slot-btn-gridlines'), 'SL'], [S('slot-field-zoom'), 'Q'], [S('slot-btn-fts'), 'W']]
    };

    // PowerPoint's backstage letters
    var BACKSTAGE = {
        'fm-btn-elx-home': 'H', 'fm-btn-create': 'N', 'fm-btn-elx-open': 'O', 'fm-btn-info': 'I', 'fm-btn-save': 'S',
        'fm-btn-save-copy': 'Y', 'fm-btn-download': 'A', 'fm-btn-print-with-preview': 'P', 'fm-btn-elx-share': 'Z',
        'fm-btn-elx-export': 'E', 'fm-btn-elx-close': 'C', 'fm-btn-settings': 'T', 'fm-btn-help': 'H1', 'fm-btn-elx-about': 'D'
    };

    PE.Controllers.PowerPointKeyTips = Common.Controllers.ElixirKeyTips.extend({
        tabs: TABS,
        commands: COMMANDS,
        backstage: BACKSTAGE
    });
});

/*
 * SPDX-FileCopyrightText: 2026 ElixirOffice contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 *
 * PE.Controllers.PowerPointScreenTips: Office ScreenTips (command name and
 * shortcut, then one line on what it does) for the presentation ribbon, on the
 * shared controller (common/main/lib/controller/ElixirScreenTips.js).
 * Translations: locale PE.Controllers.PowerPointScreenTips.*.
 */
define([
    'core',
    'common/main/lib/controller/ElixirScreenTips'
], function () {
    'use strict';

    var S = function (id) { return '#' + id; };

    var TIPS = [
        [S('slot-btn-paste'), 'descPaste'], [S('slot-btn-cut'), 'descCut'], [S('slot-btn-copy'), 'descCopy'], [S('slot-btn-copystyle'), 'descFormatPainter'],
        [S('slot-addslide'), 'descNewSlide'], [S('slot-elx-ins-newslide'), 'descNewSlide'], [S('slot-changeslide'), 'descLayout'], [S('slot-elx-reset-slide'), 'descReset'],
        [S('slot-field-fontname'), 'descFont'], [S('slot-field-fontsize'), 'descFontSize'], [S('slot-btn-incfont'), 'descGrowFont'], [S('slot-btn-decfont'), 'descShrinkFont'],
        [S('slot-btn-bold'), 'descBold'], [S('slot-btn-italic'), 'descItalic'], [S('slot-btn-underline'), 'descUnderline'], [S('slot-btn-fontcolor'), 'descFontColor'],
        [S('slot-btn-markers'), 'descBullets'], [S('slot-btn-numbering'), 'descNumbering'], [S('slot-btn-linespace'), 'descLineSpacing'],
        [S('slot-btn-arrange-shape'), 'descArrange'], [S('slot-btn-align-shape'), 'descAlign'],
        [S('slot-elx-find'), 'descFind'], [S('slot-btn-replace'), 'descReplace'], [S('slot-btn-select-all'), 'descSelectAll'],
        [S('slot-btn-inserttable'), 'descTable'], [S('slot-btn-insertchart'), 'descChart'], [S('slot-btn-inssmartart'), 'descSmartArt'],
        [S('slot-btn-instextart'), 'descWordArt'], [S('slot-btn-editheader'), 'descHeaderFooter'],
        [S('slot-btn-colorschemas'), 'descVariants'], [S('slot-btn-slidesize'), 'descSlideSize'], [S('slot-elx-format-bg'), 'descFormatBackground'],
        [S('slot-elx-show-begin'), 'descFromBeginning'], [S('slot-elx-show-current'), 'descFromCurrent'], [S('slot-elx-show-presenter'), 'descPresenter'],
        [S('slot-elx-show-hide'), 'descHideSlide'],
        ['header.btnUndo', 'descUndo'], ['header.btnRedo', 'descRedo'], ['header.btnSave', 'descSave'], ['header.btnPrint', 'descPrint']
    ];

    PE.Controllers.PowerPointScreenTips = Common.Controllers.ElixirScreenTips.extend(_.extend({
        tips: TIPS,

        // the tips go on the controls inside the slots
        resolve: function (path) {
            if (/^[#.]/.test(path)) return $('#toolbar').find(path).first().children().first();
            return Common.Controllers.ElixirScreenTips.prototype.resolve.call(this, path);
        },

        descPaste: 'Add content from the Clipboard to your slide.',
        descCut: 'Remove the selection and put it on the Clipboard.',
        descCopy: 'Put a copy of the selection on the Clipboard.',
        descFormatPainter: 'Copy the look of the selection and apply it to other content.',
        descNewSlide: 'Add a slide; choose its layout from the arrow.',
        descLayout: 'Change the arrangement of the content on the slide.',
        descReset: 'Put the placeholders of the slide back to their default position, size and formatting.',
        descFont: 'Pick a new font for your text.',
        descFontSize: 'Change the size of your text.',
        descGrowFont: 'Make your text a bit bigger.',
        descShrinkFont: 'Make your text a bit smaller.',
        descBold: 'Make your text bold.',
        descItalic: 'Italicize your text.',
        descUnderline: 'Underline your text.',
        descFontColor: 'Change the color of your text.',
        descBullets: 'Create a bulleted list.',
        descNumbering: 'Create a numbered list.',
        descLineSpacing: 'Choose the space between the lines of text.',
        descArrange: 'Bring objects forward or send them back, group them.',
        descAlign: 'Line objects up on the slide or with each other.',
        descFind: 'Find text in the presentation.',
        descReplace: 'Search for text and replace it.',
        descSelectAll: 'Select all the objects on the slide.',
        descTable: 'Insert a table to organize information.',
        descChart: 'Add a chart to show data.',
        descSmartArt: 'Insert a SmartArt graphic to communicate information visually.',
        descWordArt: 'Add decorative text to the slide.',
        descHeaderFooter: 'Add a footer, the date or the slide number.',
        descVariants: 'Choose the colors of the theme.',
        descSlideSize: 'Change the size and shape of the slides.',
        descFormatBackground: 'Change the background of the slide: color, gradient or picture.',
        descFromBeginning: 'Start the slide show from the first slide.',
        descFromCurrent: 'Start the slide show from the current slide.',
        descPresenter: 'Show your notes, a timer and the slides while you present.',
        descHideSlide: 'Skip this slide during the slide show.',
        descUndo: 'Undo your last action.',
        descRedo: 'Redo the action you undid.',
        descSave: 'Save your presentation.',
        descPrint: 'Print your presentation.'
    }, PE.Controllers.PowerPointScreenTips || {}));
});

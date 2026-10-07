/*
 * (c) Copyright Ascensio System SIA 2010-2024
 *
 * This program is a free software product. You can redistribute it and/or
 * modify it under the terms of the GNU Affero General Public License (AGPL)
 * version 3 as published by the Free Software Foundation. In accordance with
 * Section 7(a) of the GNU AGPL its Section 15 shall be amended to the effect
 * that Ascensio System SIA expressly excludes the warranty of non-infringement
 * of any third-party rights.
 *
 * This program is distributed WITHOUT ANY WARRANTY; without even the implied
 * warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR  PURPOSE. For
 * details, see the GNU AGPL at: http://www.gnu.org/licenses/agpl-3.0.html
 *
 * The  interactive user interfaces in modified source and object code versions
 * of the Program must display Appropriate Legal Notices, as required under
 * Section 5 of the GNU AGPL version 3.
 *
 * All the Product's GUI elements, including illustrations and icon sets, as
 * well as technical writing content are licensed under the terms of the
 * Creative Commons Attribution-ShareAlike 4.0 International. See the License
 * terms at http://creativecommons.org/licenses/by-sa/4.0/legalcode
 *
 */

/**
 *  ScreenTips.js
 *
 *  Word's ScreenTips (engine: common/main/lib/controller/ElixirScreenTips.js):
 *  the commands and their one-line descriptions.
 */

define([
    'core',
    'common/main/lib/controller/ElixirScreenTips'
], function () {
    'use strict';

    // accessor (see resolve) -> description string of this controller
    var TIPS = [
        ['tb.btnPaste', 'descPaste'], ['tb.btnCut', 'descCut'], ['tb.btnCopy', 'descCopy'], ['tb.btnCopyStyle', 'descFormatPainter'],
        ['tb.cmbFontName', 'descFont'], ['tb.cmbFontSize', 'descFontSize'], ['tb.btnIncFontSize', 'descGrowFont'], ['tb.btnDecFontSize', 'descShrinkFont'],
        ['tb.btnChangeCase', 'descChangeCase'], ['tb.btnClearStyle', 'descClearFormatting'],
        ['tb.btnBold', 'descBold'], ['tb.btnItalic', 'descItalic'], ['tb.btnUnderline', 'descUnderline'], ['tb.btnStrikeout', 'descStrikethrough'],
        ['tb.btnSubscript', 'descSubscript'], ['tb.btnSuperscript', 'descSuperscript'], ['tb.btnHighlightColor', 'descHighlight'], ['tb.btnFontColor', 'descFontColor'],
        ['tb.btnMarkers', 'descBullets'], ['tb.btnNumbers', 'descNumbering'], ['tb.btnMultilevels', 'descMultilevel'],
        ['tb.btnDecLeftOffset', 'descDecreaseIndent'], ['tb.btnIncLeftOffset', 'descIncreaseIndent'], ['tb.btnShowHidenChars', 'descShowHide'],
        ['tb.btnAlignLeft', 'descAlignLeft'], ['tb.btnAlignCenter', 'descCenter'], ['tb.btnAlignRight', 'descAlignRight'], ['tb.btnAlignJust', 'descJustify'],
        ['tb.btnLineSpace', 'descLineSpacing'], ['tb.btnParagraphColor', 'descShading'], ['tb.btnBorders', 'descBorders'],
        ['tb.btnFind', 'descFind'], ['tb.btnReplace', 'descReplace'], ['tb.btnSelectAll', 'descSelectAll'],
        ['tb.btnUndo', 'descUndo'], ['tb.btnRedo', 'descRedo'], ['tb.btnSave', 'descSave'], ['tb.btnPrint', 'descPrint'],
        ['header.btnUndo', 'descUndo'], ['header.btnRedo', 'descRedo'], ['header.btnSave', 'descSave'], ['header.btnPrint', 'descPrint'],
        ['tb.btnInsertTable', 'descTable'], ['#tlbtn-insertimage-0', 'descPicture'], ['tb.btnInsertShape', 'descShapes'], ['tb.btnInsertChart', 'descChart'],
        ['links.btnsHyperlink.0', 'descLink'], ['#tlbtn-addcomment-0', 'descComment'], ['tb.btnsPageBreak.0', 'descPageBreak'],
        ['#tlbtn-headerfooter-0', 'descHeader'], ['#tlbtn-insertpagenum-0', 'descPageNumber'], ['tb.btnInsertText', 'descTextBox'],
        ['tb.btnInsertEquation', 'descEquation'], ['tb.btnInsertSymbol', 'descSymbol']
    ];

    DE.Controllers.ScreenTips = Common.Controllers.ElixirScreenTips.extend(_.extend({
        tips: TIPS,

        roots: function() {
            var app = this.getApplication(),
                roots = Common.Controllers.ElixirScreenTips.prototype.roots.call(this);
            roots.links = function() { var c = app.getController('Links'); return c && c.view; };
            return roots;
        },

        descPaste: 'Add content on the Clipboard to your document.',
        descCut: 'Remove the selection and put it on the Clipboard so you can paste it somewhere else.',
        descCopy: 'Put a copy of the selection on the Clipboard so you can paste it somewhere else.',
        descFormatPainter: 'Like the look of one selection? You can apply that look to other content in the document.',
        descFont: 'Pick a new font for your text.',
        descFontSize: 'Change the size of your text.',
        descGrowFont: 'Make your text a bit bigger.',
        descShrinkFont: 'Make your text a bit smaller.',
        descChangeCase: 'Change the selected text to uppercase, lowercase, or other common capitalizations.',
        descClearFormatting: 'Remove all formatting from the selection, leaving only the normal, unformatted text.',
        descBold: 'Make your text bold.',
        descItalic: 'Italicize your text.',
        descUnderline: 'Underline your text.',
        descStrikethrough: 'Cross something out by drawing a line through it.',
        descSubscript: 'Type very small letters just below the line of text.',
        descSuperscript: 'Type very small letters just above the line of text.',
        descHighlight: 'Make text stand out by highlighting it in a bright color.',
        descFontColor: 'Change the color of your text.',
        descBullets: 'Create a bulleted list.',
        descNumbering: 'Create a numbered list.',
        descMultilevel: 'Create a multilevel list to organize items or create an outline.',
        descDecreaseIndent: 'Move your paragraph closer to the margin.',
        descIncreaseIndent: 'Move your paragraph farther away from the margin.',
        descShowHide: 'Show paragraph marks and other hidden formatting symbols.',
        descAlignLeft: 'Align your content with the left margin.',
        descCenter: 'Center your content on the page.',
        descAlignRight: 'Align your content with the right margin.',
        descJustify: 'Distribute your text evenly between the margins.',
        descLineSpacing: 'Choose how much space appears between lines of text or between paragraphs.',
        descShading: 'Color the background behind the selected text or paragraph.',
        descBorders: 'Add or remove borders from your selection.',
        descFind: 'Find text or other content in the document.',
        descReplace: 'Search for text you\'d like to change and replace it with something else.',
        descSelectAll: 'Select all the content of the document.',
        descUndo: 'Undo your last action.',
        descRedo: 'Redo the action you undid.',
        descRepeat: 'Repeat the last command or action.',
        descSave: 'Save your document.',
        descPrint: 'Print your document.',
        descTable: 'Add a table to organize information.',
        descPicture: 'Insert pictures from your computer or from the web.',
        descShapes: 'Insert ready-made shapes, such as circles, squares, and arrows.',
        descChart: 'Add a chart to illustrate and compare data.',
        descLink: 'Create a link in your document for quick access to webpages and files.',
        descComment: 'Add a note about this part of the document.',
        descPageBreak: 'End the current page here and move to the next page.',
        descHeader: 'Add a header and a footer to every page.',
        descPageNumber: 'Number the pages in your document.',
        descTextBox: 'Add a text box that you can place anywhere on the page.',
        descEquation: 'Add common mathematical equations to your document.',
        descSymbol: 'Add symbols that aren\'t on your keyboard.'

    }, DE.Controllers.ScreenTips || {}));
});

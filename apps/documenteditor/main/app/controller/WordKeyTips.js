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
 *  WordKeyTips.js
 *
 *  Word's KeyTip letters for the access keys of the ribbon (Alt, then H for
 *  Home, then 1 for Bold...). The letters are those of Word for Microsoft 365
 *  (en-US); a control without a Word equivalent keeps the letter that the hint
 *  manager assigns automatically. Level 0 shows the tabs and the quick access
 *  toolbar only, as Word does.
 */

define([
    'core'
], function () {
    'use strict';

    // tab (data-tab of the ribbon tab) -> KeyTip
    var TABS = {
        file: 'F', home: 'H', ins: 'N', draw: 'JI', design: 'G', layout: 'P', links: 'S',
        review: 'R', view: 'W', plugins: 'X', headerfooter: 'JH', forms: 'L', elixir: 'A'
    };

    // per tab: [accessor, KeyTip]; an accessor is "<controller>.<path>" (see resolve) or a CSS selector
    var COMMANDS = {
        home: [
            ['tb.btnPaste', 'V'], ['tb.btnCut', 'X'], ['tb.btnCopy', 'C'], ['tb.btnCopyStyle', 'FP'],
            ['tb.cmbFontName', 'FF'], ['tb.cmbFontSize', 'FS'], ['tb.btnIncFontSize', 'FG'], ['tb.btnDecFontSize', 'FK'],
            ['tb.btnChangeCase', '7'], ['tb.btnClearStyle', 'E'],
            ['tb.btnBold', '1'], ['tb.btnItalic', '2'], ['tb.btnUnderline', '3'], ['tb.btnStrikeout', '4'],
            ['tb.btnSubscript', '5'], ['tb.btnSuperscript', '6'], ['tb.btnHighlightColor', 'I'], ['tb.btnFontColor', 'FC'],
            ['tb.btnMarkers', 'U'], ['tb.btnNumbers', 'N'], ['tb.btnMultilevels', 'M'],
            ['tb.btnDecLeftOffset', 'AO'], ['tb.btnIncLeftOffset', 'AI'], ['tb.btnShowHidenChars', '8'],
            ['tb.btnAlignLeft', 'AL'], ['tb.btnAlignCenter', 'AC'], ['tb.btnAlignRight', 'AR'], ['tb.btnAlignJust', 'AJ'],
            ['tb.btnLineSpace', 'K'], ['tb.btnParagraphColor', 'H'], ['tb.btnBorders', 'B'],
            ['tb.listStyles', 'L'], ['tb.btnFind', 'FD'], ['tb.btnReplace', 'R'], ['tb.btnSelectAll', 'SL'],
            ['.group-launcher[data-launcher="font"]', 'FN'], ['.group-launcher[data-launcher="paragraph"]', 'PG']
        ],
        ins: [
            ['tb.btnBlankPage', 'NP'], ['tb.btnsPageBreak.0', 'B'], ['tb.btnInsertTable', 'T'], ['tb.btnsInsertImage.0', 'P'],
            ['tb.btnInsertShape', 'SH'], ['tb.btnInsertSmartArt', 'M'], ['tb.btnInsertChart', 'C'],
            ['links.btnsHyperlink.0', 'I'], ['tb.btnsAddComment.0', 'L'], ['tb.btnsHeadersFooters.0', 'H'],
            ['tb.btnsPageNumber.0', 'NU'], ['tb.btnInsertText', 'X'], ['tb.btnInsertTextArt', 'W'], ['tb.btnDropCap', 'RC'],
            ['tb.btnsDateTime.0', 'D'], ['tb.btnTextFromFile', 'J'], ['tb.btnsInsertField.0', 'Q'],
            ['tb.btnInsertEquation', 'E'], ['tb.btnInsertSymbol', 'U'],
            ['#tlbtn-insertimage-0', 'P'], ['#tlbtn-addcomment-0', 'L'], ['#tlbtn-headerfooter-0', 'H'],
            ['#tlbtn-insertpagenum-0', 'NU'], ['#tlbtn-insertdatetime-0', 'D'], ['#tlbtn-insertfield-0', 'Q']
        ],
        design: [
            ['tb.btnColorSchemas', 'TC'], ['tb.btnWatermark', 'PW'], ['tb.btnPageColor', 'PC']
        ],
        layout: [
            ['tb.btnPageMargins', 'M'], ['tb.btnPageOrient', 'O'], ['tb.btnPageSize', 'SZ'], ['tb.btnColumns', 'J'],
            ['tb.btnsPageBreak.1', 'B'], ['tb.btnLineNumbers', 'LN'], ['tb.btnHyphenation', 'H'],
            ['#slot-spin-ind-left', 'IL'], ['#slot-spin-ind-right', 'IR'], ['#slot-spin-space-before', 'SB'], ['#slot-spin-space-after', 'SA'],
            ['tb.btnImgWrapping', 'TW'], ['tb.btnImgForward', 'AF'], ['tb.btnImgBackward', 'AE'], ['tb.btnImgAlign', 'AA'], ['tb.btnImgGroup', 'G'],
            ['.group-launcher[data-launcher="pagesetup"]', 'SP']
        ],
        links: [
            ['links.btnsContents.0', 'T'], ['links.btnAddText', 'A'], ['links.btnContentsUpdate', 'U'], ['links.btnsNotes.0', 'F'],
            ['links.btnBookmarks', 'K'], ['links.btnCaption', 'P'], ['links.btnCrossRef', 'RF'], ['links.btnTableFigures', 'G'],
            ['links.btnTableFiguresUpdate', 'V'], ['links.btnsHyperlink.1', 'I']
        ],
        review: [
            ['review.btnsAddComment.0', 'C'], ['#tlbtn-addcomment-1', 'C'], ['review.btnCommentRemove', 'D'], ['review.btnTurnOn', 'G'],
            ['review.btnReviewView', 'TD'], ['review.btnAccept', 'A2'], ['review.btnReject', 'J'], ['review.btnCompare', 'M'],
            ['review.btnHistory', 'HI']
        ],
        view: [
            ['viewtab.btnNavigation', 'K'], ['viewtab.btnZoom100', 'J'], ['viewtab.btnMultiplePages', 'M'],
            ['#slot-chk-rulers', 'R'], ['viewtab.btnMacros', 'H'], ['viewtab.btnInterfaceTheme', 'TH']
        ]
    };

    // Backstage (File): Word's letters
    var BACKSTAGE = {
        'fm-btn-create': 'N', 'fm-btn-recent': 'O', 'fm-btn-info': 'I', 'fm-btn-save': 'S', 'fm-btn-save-copy': 'A',
        'fm-btn-print': 'P', 'fm-btn-download': 'E', 'fm-btn-rename': 'M', 'fm-btn-protect': 'R',
        'fm-btn-history': 'V', 'fm-btn-back': 'C', 'fm-btn-settings': 'T', 'fm-btn-help': 'H'
    };

    DE.Controllers.WordKeyTips = Backbone.Controller.extend({
        initialize: function() {
            Common.NotificationCenter.on('hints:before-show', _.bind(this.onBeforeShow, this));
            Common.NotificationCenter.on('app:ready', _.bind(this.onAppReady, this));
        },

        onAppReady: function() {
            var me = this;
            // level 0: the ribbon tabs and the quick access toolbar (not the side panes and the status bar)
            Common.UI.HintManager.setItemFilter && Common.UI.HintManager.setItemFilter(function(item, level) {
                if (level !== 0) return true;
                return item.closest('#toolbar, #header, .hedset, #app-title, #box-document-title').length > 0;
            });
            // Alt+letter held together opens the hints of a tab (Alt+F stays File, handled by the left menu)
            Common.UI.HintManager.setChordLetters && Common.UI.HintManager.setChordLetters(function() {
                // (not X: Alt+X converts Unicode codes; not on macOS, where Option+letter types characters)
                if (Common.Utils.isMac) return '';
                return _.filter(_.values(TABS), function(t) { return t.length === 1 && 'FAX'.indexOf(t) < 0; }).join('');
            });
            me.ready = true;
        },

        resolve: function(path) {
            var app = this.getApplication();
            var roots = {
                tb: function() { var c = app.getController('Toolbar'); return c && c.toolbar; },
                links: function() { var c = app.getController('Links'); return c && c.view; },
                review: function() { var c = app.getController('Common.Controllers.ReviewChanges'); return c && c.view; },
                viewtab: function() { var c = app.getController('ViewTab'); return c && c.view; },
                header: function() { var c = app.getController('Viewport'); return c && c.header; }
            };
            if (/^[#.\[]/.test(path))
                return $('#toolbar').find(path).first();
            var parts = path.split('.'), root = roots[parts[0]] && roots[parts[0]]();
            for (var i = 1; root && i < parts.length; i++)
                root = root[parts[i]];
            if (!root) return $();
            var $el = root.cmpEl || root.$el || (root.jquery ? root : null);
            return $el ? $el : $();
        },

        hintElement: function($el) {
            if (!$el || !$el.length) return $();
            if ($el.is('[data-hint]')) return $el;
            return $el.find('[data-hint]').first();
        },

        setTip: function($el, tip) {
            var $hint = this.hintElement($el);
            if (!$hint.length) return;
            $hint.attr('data-hint-title', tip).removeAttr('data-hint-title-lang');
        },

        onBeforeShow: function(level) {
            if (!this.ready) return;
            var me = this;
            if (level === 0) {
                $('#toolbar .ribtab > a[data-tab]').each(function() {
                    var tip = TABS[$(this).attr('data-tab')];
                    tip && $(this).attr('data-hint-title', tip).removeAttr('data-hint-title-lang');
                });
                // quick access toolbar: numbers from left to right, as in Word
                var n = 1;
                $('[data-hint="0"]').filter(':visible').filter(function() {
                    return $(this).closest('#toolbar').length === 0 && $(this).closest('#header, .hedset, #app-title, #box-document-title').length > 0 &&
                        /^(S|P|Z|Y|\d)$/.test($(this).attr('data-hint-title') || '');
                }).sort(function(a, b) { return a.getBoundingClientRect().left - b.getBoundingClientRect().left; }).each(function() {
                    // Word draws the quick access KeyTips under the buttons
                    $(this).attr({'data-hint-title': String(n++), 'data-hint-direction': 'bottom', 'data-hint-offset': '-6, 0'}).removeAttr('data-hint-title-lang');
                });
                // Search (Word: Alt, Q)
                me.setTip(me.resolve('header.btnSearch'), 'Q');
                return;
            }
            if ($('#file-menu-panel').is(':visible')) {
                _.each(BACKSTAGE, function(tip, id) {
                    var $el = $('#' + id);
                    me.setTip($el, tip);
                });
                return;
            }
            var tab = $('#toolbar .ribtab.active > a[data-tab]').attr('data-tab');
            _.each(COMMANDS[tab] || [], function(entry) {
                me.setTip(me.resolve(entry[0]), entry[1]);
            });
        }
    });
});

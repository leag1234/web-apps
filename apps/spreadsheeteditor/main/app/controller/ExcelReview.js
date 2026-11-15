/*
 * SPDX-FileCopyrightText: 2026 ElixirOffice contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 *
 * SSE.Controllers.ExcelReview: what the Review tab of the spreadsheet editor does beyond the stock controls
 * (the buttons are created by SSE.Controllers.RibbonTabs).
 *
 *   Thesaurus           synonyms of the text of the active cell, from the offline MyThes files
 *                       (Common.util.Thesaurus); Insert replaces the content of the cell
 *   Workbook Statistics worksheets, cells with data, formulas, tables, PivotTables, charts, pictures,
 *                       shapes and comments of the whole workbook
 *   Previous / Next     the comments of the active worksheet, row by row, wrapping around; the commented cell
 *                       is selected and scrolled into view, and its note is shown
 */

define([
    'core',
    'common/main/lib/util/Thesaurus',
    'common/main/lib/component/Window'
], function () {
    'use strict';

    SSE.Controllers.ExcelReview = Backbone.Controller.extend(_.extend({
        initialize: function () {
            this._lastComment = null;
        },

        getApi: function () {
            if (!this.api) {
                var vp = this.getApplication().getController('Viewport');
                this.api = vp && vp.getApi && vp.getApi();
            }
            return this.api;
        },

        // ---- comments ------------------------------------------------------------------------------
        // the visible comments of the active worksheet in reading order: [{id, row, col}]
        commentList: function (api, collection) {
            var list = [];
            collection.each(function (m) {
                if (m.get('hide') || m.get('removed')) return;
                var c = api.asc_findComment(m.get('uid'));
                c && list.push({id: m.get('uid'), row: c.asc_getRow(), col: c.asc_getCol()});
            });
            return _.sortBy(list, function (c) { return c.row * 100000 + c.col; });
        },

        // A1 (or R1C1, when the workbook shows it) reference of a 0-based row / column
        cellRef: function (row, col) {
            var r1c1 = false;
            try { r1c1 = !!(window.AscCommonExcel && AscCommonExcel.g_R1C1Mode); } catch (e) {}
            if (r1c1) return 'R' + (row + 1) + 'C' + (col + 1);
            var letters = '', n = col + 1;
            while (n > 0) { letters = String.fromCharCode(65 + (n - 1) % 26) + letters; n = Math.floor((n - 1) / 26); }
            return letters + (row + 1);
        },

        gotoComment: function (dir) {
            var api = this.getApi(), cc = this.getApplication().getController('Common.Controllers.Comments');
            if (!api || !cc || !cc.collection) return;
            var list = this.commentList(api, cc.collection);
            if (!list.length) return;
            var cur = this._lastComment ? _.findIndex(list, function (c) { return c.id === this._lastComment; }, this) : -1;
            var next = cur < 0 ? (dir > 0 ? 0 : list.length - 1) : (cur + dir + list.length) % list.length;
            this._lastComment = list[next].id;
            // like Excel: the commented cell becomes the active cell (and is scrolled into view), then its note is shown
            api.asc_findCell && api.asc_findCell(this.cellRef(list[next].row, list[next].col));
            api.asc_selectComment(this._lastComment);
            api.asc_showComment(this._lastComment);
        },

        // ---- Workbook Statistics -------------------------------------------------------------------
        // numbers of the whole workbook, from the workbook model (wb: Asc.editor.wbModel)
        collectStats: function (wb) {
            var r = {sheets: 0, hidden: 0, cells: 0, formulas: 0, tables: 0, pivots: 0, charts: 0, pictures: 0, shapes: 0, comments: 0},
                T = typeof AscDFH !== 'undefined' ? AscDFH : {};
            _.each(wb.aWorksheets || [], function (ws) {
                r.sheets++;
                ws.getHidden() && r.hidden++;
                var rows = ws.getRowsCount(), cols = ws.getColsCount();
                if (rows > 0 && cols > 0) {
                    ws.getRange3(0, 0, rows - 1, cols - 1)._foreachNoEmpty(function (cell) {
                        if (cell.isNullText()) return;
                        r.cells++;
                        cell.formulaParsed && r.formulas++;
                    });
                }
                r.tables += (ws.TableParts || []).length;
                r.pivots += (ws.pivotTables || []).length;
                r.comments += (ws.aComments || []).length;
                _.each(ws.Drawings || [], function (d) {
                    var g = d && d.graphicObject, type = g && g.getObjectType && g.getObjectType();
                    if (type !== undefined && type === T.historyitem_type_ChartSpace) r.charts++;
                    else if (type !== undefined && type === T.historyitem_type_ImageShape) r.pictures++;
                    else r.shapes++;
                });
            });
            return r;
        },

        showStatistics: function () {
            var api = this.getApi(), v = this.getApplication().getController('RibbonTabs').view;
            if (!api || !api.wbModel) return;
            var stats = this.collectStats(api.wbModel),
                rows = [['sheets', v.txtStatSheets], ['hidden', v.txtStatHidden], ['cells', v.txtStatCells], ['formulas', v.txtStatFormulas],
                        ['tables', v.txtStatTables], ['pivots', v.txtStatPivots], ['charts', v.txtStatCharts], ['pictures', v.txtStatPictures],
                        ['shapes', v.txtStatShapes], ['comments', v.txtStatComments]],
                html = '<div class="elx-word-count"><div class="elx-wc-title">' + v.txtStatistics + '</div><table>' +
                    _.map(rows, function (x) { return '<tr><td>' + x[1] + '</td><td class="elx-wc-val">' + stats[x[0]].toLocaleString() + '</td></tr>'; }).join('') +
                    '</table></div>';
            var win = new Common.UI.Window({title: v.txtWorkbookStatsTitle, width: 320, cls: 'modal-dlg', buttons: ['close'], tpl: html});
            win.show();
            // a plain window does not bind its footer button
            win.$window.find('.dlg-btn').on('click', function () { win.close(); });
        },

        // ---- Thesaurus -----------------------------------------------------------------------------
        // the language of the text: the spell checking language of the workbook, else the one of the interface
        thesaurusLang: function () {
            var s = Common.Utils.InternalSettings.get('sse-spellcheck-locale'), mode = this.getApplication().getController('Main').appOptions || {};
            if (s) return parseInt(s);
            return mode.lang ? parseInt(Common.util.LanguageInfo.getLocalLanguageCode(mode.lang)) : 0x0409;
        },

        cellWord: function (api) {
            var info = api.asc_getCellInfo(), text = String((info && info.asc_getText && info.asc_getText()) || '');
            return text.replace(/[\r\n]+/g, ' ').trim().split(/\s+/).slice(0, 3).join(' ');
        },

        openThesaurus: function () {
            var me = this, api = this.getApi(), v = this.getApplication().getController('RibbonTabs').view, Th = Common.util.Thesaurus;
            if (!api) return;
            if (this._thesWin) { this._thesWin.close(); }
            var canEdit = !!(this.getApplication().getController('Main').appOptions || {}).isEdit,
                lcid = this.thesaurusLang(),
                tpl = '<div class="elx-sse-thes"><div class="elx-th-row"><input type="text" class="form-control elx-th-input"><button type="button" class="btn btn-text-default elx-th-go"></button></div><div class="elx-th-result"></div></div>',
                win = this._thesWin = new Common.UI.Window({title: v.txtThesaurusTitle, width: 340, height: 360, cls: 'modal-dlg', modal: false, buttons: ['close'], tpl: tpl});
            win.show();
            win.$window.find('.dlg-btn').on('click', function () { win.close(); });
            win.on('close', function () { me._thesWin = null; api.asc_enableKeyEvents(true); });
            var $w = win.$window, $input = $w.find('.elx-th-input'), $res = $w.find('.elx-th-result');
            $w.find('.elx-th-go').text(v.txtLookUp);
            $input.attr('aria-label', v.txtSearchFor);
            var show = function (word) {
                $input.val(word);
                if (!word) { $res.text(v.txtTypeWord); return; }
                $res.text(v.txtSearching);
                Th.load(lcid).then(function (thes) {
                    if (!thes) { $res.text(v.txtNoThesaurus); return null; }
                    return Th.lookup(thes, word).then(function (meanings) {
                        $res.empty();
                        if (!meanings || !meanings.length) { $res.text(Common.Utils.String.format(v.txtNoResults, word)); return; }
                        _.each(meanings, function (m) {
                            $('<div class="elx-th-pos"></div>').text(m.pos ? m.pos : word).appendTo($res);
                            _.each(m.words, function (syn) {
                                var $row = $('<div class="elx-th-syn"><span class="elx-th-name"></span><button type="button" class="elx-th-insert"></button></div>');
                                $row.find('.elx-th-name').text(syn);
                                canEdit && $row.find('.elx-th-insert').text('↵').attr('title', v.txtInsert);
                                !canEdit && $row.find('.elx-th-insert').hide();
                                $row.on('click', function (e) {
                                    if ($(e.target).closest('.elx-th-insert').length) me.insertSynonym(syn);
                                    else show(syn.replace(/\s*\(.*\)$/, ''));
                                });
                                $res.append($row);
                            });
                        });
                    });
                });
            };
            $input.on('focus', function () { api.asc_enableKeyEvents(false); })
                .on('keydown', function (e) { if (e.keyCode === Common.UI.Keys.RETURN) { e.preventDefault(); show($input.val().trim()); } });
            $w.find('.elx-th-go').on('click', function () { show($input.val().trim()); });
            show(this.cellWord(api));
        },

        // replaces the content of the active cell by the synonym
        insertSynonym: function (syn) {
            var api = this.getApi();
            if (!api) return;
            api.asc_enableKeyEvents(true);
            api.asc_PasteData(AscCommon.c_oAscClipboardDataFormat.Text, syn.replace(/\s*\(.*\)$/, ''));
            Common.NotificationCenter.trigger('edit:complete');
        }
    }, SSE.Controllers.ExcelReview || {}));
});

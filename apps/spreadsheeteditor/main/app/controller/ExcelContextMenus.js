/*
 * SPDX-FileCopyrightText: 2026 ElixirOffice contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 *
 * SSE.Controllers.ExcelContextMenus: the context menus of the cells, rows,
 * columns and sheet tabs in Excel's order, with Excel's words and the
 * shortcuts on the right.
 *
 * The stock code addresses some menu items by their index, so the items array
 * is kept: only the DOM order changes, each time the menu opens (the stock
 * code has set the visibility of the items by then). The stock separators are
 * hidden; ours are shown between two visible groups.
 */

define([
    'core'
], function () {
    'use strict';

    // cell / row / column / select-all menu (DocumentHolder ssMenu): groups of view properties
    var CELL_GROUPS = [
        ['pmiCut', 'pmiCopy', 'pmiPaste'],
        ['pmiSelectTable', 'pmiInsertEntire', 'pmiInsertCells', 'pmiInsertTable', 'pmiDeleteEntire', 'pmiDeleteCells', 'pmiDeleteTable', 'pmiClear'],
        ['pmiSparklines'],
        ['pmiFilterCells', 'pmiSortCells', 'pmiReapply'],
        ['mnuRefreshPivot', 'mnuPivotSort', 'mnuPivotFilter', 'mnuSubtotalField', 'mnuExpandCollapsePivot', 'mnuGroupPivot', 'mnuUnGroupPivot',
            'mnuDeleteField', 'mnuSummarize', 'mnuShowAs', 'mnuShowDetails', 'mnuFieldSettings', 'mnuPivotSettings'],
        ['pmiAddComment'],
        ['pmiCellFormat', 'pmiEntriesList', 'pmiAddNamedRange', 'menuAddHyperlink', 'menuHyperlink', 'pmiGetRangeList'],
        ['pmiNumFormat', 'pmiCondFormat', 'pmiInsFunction'],
        ['pmiRowHeight', 'pmiColumnWidth', 'pmiEntireHide', 'pmiEntireShow'],
        ['pmiFreezePanes'],
        ['pmiAssistant']
    ];

    // Excel's captions (keys of this controller's strings)
    var CELL_LOOK = {
        pmiCellFormat: 'textFormatCells', pmiAddNamedRange: 'textDefineName',
        pmiEntriesList: 'textPickFromList', pmiAddComment: 'textNewComment', menuAddHyperlink: 'textLink',
        pmiEntireShow: 'textUnhide'
    };

    // shortcuts on the right: [Windows/Linux, macOS]
    var CELL_KEYS = {
        pmiCut: ['Ctrl+X', '⌘X'], pmiCopy: ['Ctrl+C', '⌘C'], pmiPaste: ['Ctrl+V', '⌘V'],
        pmiCellFormat: ['Ctrl+1', '⌘1'], menuAddHyperlink: ['Ctrl+K', '⌘K'], pmiAddComment: ['Ctrl+Shift+F2', '⇧⌘F2'],
        pmiInsFunction: ['Shift+F3', '⇧F3']
    };

    // sheet tab menu (Statusbar tabMenu): Excel's order, by item value (Unhide has none: index 5)
    var TAB_ORDER = ['ins', 'del', 'ren', 'move-copy', 'protect', '@color', 'hide', '@unhide', '|', 'selectall', 'noselect'];

    SSE.Controllers.ExcelContextMenus = Backbone.Controller.extend(_.extend({
        initialize: function () {
            Common.NotificationCenter.on('app:ready', _.bind(this.onAppReady, this));
        },

        onAppReady: function (mode) {
            this.mode = mode || {};
            var me = this, tries = 0;
            var wait = function () {
                var dh = me.getApplication().getController('DocumentHolder'), view = dh && dh.documentHolder,
                    sb = me.getApplication().getController('Statusbar'), sbv = sb && sb.statusbar;
                if (view && view.ssMenu && sbv && sbv.tabMenu) {
                    me.setupCellMenu(view);
                    me.setupTabMenu(sbv);
                } else if (++tries < 60) setTimeout(wait, 500);
            };
            wait();
        },

        // ---- cells, rows, columns -----------------------------------------------------
        setupCellMenu: function (view) {
            var me = this, menu = view.ssMenu;
            this.view = view;
            var groups = _.map(CELL_GROUPS, function (g) { return _.compact(_.map(g, function (n) { var it = view[n]; if (it) it._excelKey = n; return it; })); });
            var placed = _.flatten(groups);
            // Excel's words, before the first rendering when possible
            _.each(CELL_LOOK, function (key, name) {
                var it = view[name];
                if (it && me[key]) { it.caption = me[key]; it.options && (it.options.caption = me[key]); it.rendered && it.setCaption(me[key]); }
            });
            menu.on('show:before', function (m, e) {
                // 'show.bs.dropdown' bubbles up from a submenu item that opens: arranging then moves the item
                // under the pointer, which opens it again (endless loop; clicks on submenu items are lost)
                if (me.fromSubmenu(e)) return;
                try { me.arrange(menu, groups, placed, CELL_KEYS); } catch (err) { console.warn('ExcelContextMenus', err); }
            });
        },

        // is this show:before event the one of a submenu item rather than of the menu itself?
        fromSubmenu: function (e) {
            return !!(e && e.target && e.target.tagName === 'LI');
        },

        visible: function (it) {
            return !!it && it.caption !== '--' && (it.rendered && it.el ? it.el.style.display !== 'none' : it.visible !== false);
        },

        arrange: function (menu, groups, placed, keys) {
            var $ul = menu.cmpEl && (menu.cmpEl.is('ul') ? menu.cmpEl : menu.cmpEl.find('> ul.dropdown-menu, ul.dropdown-menu').first());
            if (!$ul || !$ul.length) return;
            var me = this, mac = Common.Utils.isMac;
            // stock separators hidden, ours between the visible groups
            _.each(menu.items, function (it) { it.caption === '--' && it.cmpEl && it.cmpEl.css('display', 'none'); });
            $ul.children('li.elx-sep').remove();
            var others = _.filter(menu.items, function (it) { return placed.indexOf(it) < 0 && it.caption !== '--'; });
            var first = true;
            _.each(groups.concat([others]), function (items) {
                var shown = _.filter(items, function (it) { return me.visible(it); });
                if (!shown.length) {
                    _.each(items, function (it) { it.cmpEl && $ul.append(it.cmpEl); });
                    return;
                }
                if (!first) $ul.append('<li class="divider elx-sep"></li>');
                first = false;
                _.each(items, function (it) { it.cmpEl && $ul.append(it.cmpEl); });
            });
            // stock separators at the end, hidden
            _.each(menu.items, function (it) { it.caption === '--' && it.cmpEl && $ul.append(it.cmpEl); });
            // shortcuts
            _.each(menu.items, function (it) {
                var k = keys[it._excelKey];
                if (!k || !it.cmpEl) return;
                var $a = it.cmpEl.find('> a').first();
                $a.length && !$a.find('.menu-item-shortcut').length && $a.append('<span class="menu-item-shortcut">' + k[mac ? 1 : 0] + '</span>');
            });
        },

        // ---- sheet tabs -----------------------------------------------------------------
        setupTabMenu: function (sb) {
            var me = this, menu = sb.tabMenu;
            var reorder = function () {
                var $ul = menu.cmpEl && (menu.cmpEl.is('ul') ? menu.cmpEl : menu.cmpEl.find('ul.dropdown-menu').first());
                if (!$ul || !$ul.length) return;
                var byValue = {}, seps = [];
                _.each(menu.items, function (it, i) {
                    if (it.caption === '--') { seps.push(it); return; }
                    if (it.value !== undefined && it.value !== null && it.value !== '') byValue[it.value] = it;
                    else if (i === 5) byValue['@unhide'] = it;
                    else if (it.menu && !byValue['@color'] && it !== menu.items[5]) byValue['@color'] = it;
                });
                var used = [];
                _.each(TAB_ORDER, function (key) {
                    if (key === '|') {
                        var s = seps.shift();
                        s && s.cmpEl && $ul.append(s.cmpEl);
                        s && used.push(s);
                        return;
                    }
                    var it = byValue[key];
                    if (it && it.cmpEl) { $ul.append(it.cmpEl); used.push(it); }
                });
                _.each(menu.items, function (it) {
                    if (used.indexOf(it) > -1 || !it.cmpEl) return;
                    it.caption === '--' ? it.cmpEl.css('display', 'none') : $ul.append(it.cmpEl);
                });
                // Excel's words
                byValue['@unhide'] && byValue['@unhide'].setCaption(me.textUnhideSheet);
            };
            menu.on('show:before', function (m, e) {
                if (me.fromSubmenu(e)) return;    // see setupCellMenu
                try { reorder(); } catch (err) { console.warn('ExcelContextMenus', err); }
            });
        },

        textFormatCells: 'Format Cells...',
        textDefineName: 'Define Name...',
        textPickFromList: 'Pick From Drop-down List...',
        textNewComment: 'New Comment',
        textLink: 'Link',
        textUnhide: 'Unhide',
        textUnhideSheet: 'Unhide'
    }, SSE.Controllers.ExcelContextMenus || {}));
});

/*
 * SPDX-FileCopyrightText: 2026 ElixirOffice contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 *
 * SSE.Controllers.ExcelShortcuts: Excel's keyboard shortcuts that the
 * spreadsheet editor lacked or bound to something else (matrix:
 * docs/shortcuts-excel.csv of the ElixirOffice repository, from Microsoft's
 * "Keyboard shortcuts in Excel"; suite: tests/shortcuts-excel).
 *
 * The keys are caught on the window in the capture phase, before the editor's
 * own handlers, so that Excel's meaning wins where the editor had another one
 * (Ctrl+0, Ctrl+-, Ctrl+Shift+_, Alt+Shift+arrows...). Every command goes
 * through the existing API or controller (same locks and undo).
 */

define([
    'core'
], function () {
    'use strict';

    SSE.Controllers.ExcelShortcuts = Backbone.Controller.extend({
        initialize: function () {
            Common.NotificationCenter.on('app:ready', _.bind(this.onAppReady, this));
        },

        ctrl: function (n) { return this.getApplication().getController(n); },

        onAppReady: function (mode) {
            this.mode = mode || {};
            if (this.mode.isEditDiagram || this.mode.isEditMailMerge || this.mode.isEditOle || this.ready) return;
            this.ready = true;
            this.api = this.ctrl('Viewport').getApi();
            this.tbc = this.ctrl('Toolbar');
            this.tb = this.tbc && this.tbc.toolbar;
            this.mac = Common.Utils.isMac;
            this.keys = this.table();
            window.addEventListener('keydown', _.bind(this.onKeyDown, this), true);
        },

        // [key, ctrl/cmd, alt/option, shift, handler, opts]; on macOS "ctrl" is Cmd unless opts.macCtrl
        table: function () {
            var me = this, mac = this.mac, edit = !!this.mode.isEdit;
            var t = [];
            var add = function (key, c, a, s, fn, o) { t.push({key: key, c: c, a: a, s: s, fn: fn, o: o || {}}); };
            // File
            add('KeyO', 1, 0, 0, function () { return me.backstage('elx-open'); });
            add('F12', 1, 0, 0, function () { return me.backstage('elx-open'); });
            add('F12', 0, 0, 0, function () { return me.backstage('saveas'); });
            // Interface
            add('KeyU', 1, 0, 1, function () { return me.toggleFormulaBar(); }, {inEdit: true});
            add('F6', 0, 0, 0, function () { return me.cycleRegions(1); }, {anywhere: true});
            add('F6', 0, 0, 1, function () { return me.cycleRegions(-1); }, {anywhere: true});
            if (mac) {
                add('KeyO', 1, 0, 0, function () { return me.backstage('elx-open'); }, {macCtrl: true});
                add('KeyP', 1, 0, 1, function () { return me.backstage('new'); });
                add('F2', 0, 1, 0, function () { return me.backstage('saveas'); });
                add('KeyU', 1, 0, 1, function () { return me.toggleFormulaBar(); }, {macCtrl: true, inEdit: true});
            }
            !mac && add('Equal', 1, 1, 0, function () { return me.zoom(0.1); });
            !mac && add('Minus', 1, 1, 0, function () { return me.zoom(-0.1); });
            // Navigation
            add('F5', 0, 0, 0, function () { return me.goTo(); });
            add('Space', 1, 0, 1, function () { return me.selectAll(); }, {force: true});
            add('KeyG', 1, 0, 0, function () { return me.goTo(); }, {macCtrl: true});
            add('F2', 1, 0, 0, function () { return me.formulaBar(); }, {win: true});
            if (edit) {
                // Edit
                add('KeyD', 1, 0, 0, function () { return me.fill(Asc.c_oAscFillType.fillDown); });
                add('KeyR', 1, 0, 0, function () { return me.fill(Asc.c_oAscFillType.fillRight); });
                mac && add('KeyD', 1, 0, 0, function () { return me.fill(Asc.c_oAscFillType.fillDown); }, {macCtrl: true});
                mac && add('KeyR', 1, 0, 0, function () { return me.fill(Asc.c_oAscFillType.fillRight); }, {macCtrl: true});
                add('Digit9', 1, 0, 0, function () { return me.hide('rows', true); }, {macCtrl: true});
                add('Digit9', 1, 0, 1, function () { return me.hide('rows', false); }, {macCtrl: true});
                add('Digit0', 1, 0, 0, function () { return me.hide('cols', true); }, {macCtrl: true});
                add('Digit0', 1, 0, 1, function () { return me.hide('cols', false); }, {macCtrl: true});
                if (mac) {
                    add('KeyK', 1, 0, 1, function () { return me.group(true); });
                    add('KeyJ', 1, 0, 1, function () { return me.group(false); });
                } else {
                    add('ArrowRight', 0, 1, 1, function () { return me.group(true); });
                    add('ArrowLeft', 0, 1, 1, function () { return me.group(false); });
                }
                add('F11', 0, 0, 1, function () { return me.newSheet(); });
                !mac && add('F1', 0, 1, 1, function () { return me.newSheet(); });
                add('KeyT', 1, 0, 0, function () { return me.table_(); });
                mac && add('KeyT', 1, 0, 0, function () { return me.table_(); }, {macCtrl: true});
                add('F7', 0, 0, 0, function () { return me.spelling(); });
                add('F2', 0, 0, 1, function () { return me.comment(); });
                add('F2', 1, 0, 1, function () { return me.comment(); });
                add('Minus', 1, 0, 0, function () { return me.deleteCells(); });
                mac && add('Minus', 1, 0, 0, function () { return me.deleteCells(); }, {macCtrl: true});
                // Format
                !mac && add('Digit2', 1, 0, 0, function () { return me.click('btnBold', true); });
                !mac && add('Digit3', 1, 0, 0, function () { return me.click('btnItalic', true); });
                !mac && add('Digit4', 1, 0, 0, function () { return me.click('btnUnderline', true); });
                mac && add('KeyX', 1, 0, 1, function () { return me.click('btnStrikeout', true); });
                mac && add('KeyE', 1, 0, 0, function () { return me.click('btnAlignCenter', true); });
                mac && add('KeyL', 1, 0, 0, function () { return me.click('btnAlignLeft', true); });
                mac && add('Digit9', 1, 0, 1, function () { return me.hide('rows', true); });
                mac && add('Digit0', 1, 0, 1, function () { return me.hide('cols', true); });
                mac && add('ArrowRight', 1, 1, 0, function () { return me.borders(Asc.c_oAscBorderOptions.Right); });
                mac && add('ArrowLeft', 1, 1, 0, function () { return me.borders(Asc.c_oAscBorderOptions.Left); });
                mac && add('ArrowUp', 1, 1, 0, function () { return me.borders(Asc.c_oAscBorderOptions.Top); });
                mac && add('ArrowDown', 1, 1, 0, function () { return me.borders(Asc.c_oAscBorderOptions.Bottom); });
                mac && add('KeyL', 1, 0, 0, function () { return me.names('new'); }, {macCtrl: true});
                mac && add('F3', 1, 0, 0, function () { return me.names('new'); });
                add('F3', 0, 0, 0, function () { return me.names('paste'); });
                add('F1', 0, 1, 0, function () { return me.chart(); });
                mac && add('Period', 1, 0, 1, function () { return me.click('btnIncFontSize'); });
                mac && add('Comma', 1, 0, 1, function () { return me.click('btnDecFontSize'); });
                add('Digit7', 1, 0, 1, function () { return me.borders('outer'); });
                add('Minus', 1, 0, 1, function () { return me.borders('none'); });
                mac && add('Digit0', 1, 1, 0, function () { return me.borders('outer'); });
                mac && add('Minus', 1, 1, 0, function () { return me.borders('none'); });
            }
            // Formulas
            add('F9', 1, 1, 0, function () { return me.calculate(); }, {macCtrl: true});
            add('F9', 1, 1, 1, function () { return me.calculate(); }, {macCtrl: true});
            mac && add('Equal', 1, 0, 0, function () { return me.calculate(); });
            // Macros
            !mac && add('F8', 0, 1, 0, function () { return me.macros(); });
            add('F11', 0, 1, 0, function () { return me.macros(); });
            mac && add('F8', 0, 1, 0, function () { return me.macros(); });
            // macOS: redo, Find, Replace, Sort, Filter, AutoSum
            if (mac) {
                add('KeyZ', 1, 0, 1, function () { return me.click('btnRedo'); });
                add('F5', 0, 0, 1, function () { return me.find(false); });
                add('F5', 1, 0, 0, function () { return me.find(false); }, {macCtrl: true});
                add('KeyH', 1, 0, 1, function () { return me.find(true); });
                edit && add('KeyF', 1, 0, 1, function () { return me.filter(); });
                edit && add('KeyR', 1, 0, 1, function () { return me.customSort(); });
                edit && add('KeyT', 1, 0, 1, function () { return me.autoSum(); });
                add('KeyR', 1, 1, 0, function () { return me.toggleRibbon(); });
            }
            return t;
        },

        // ---- dispatch ---------------------------------------------------------------------
        onKeyDown: function (e) {
            if (!this.api) return;
            if (this.pasteSpecialKey(e)) return;
            var mac = this.mac, me = this;
            var entry = _.find(this.keys, function (k) {
                if (k.key !== e.code) return false;
                var c = mac ? (k.o.macCtrl ? e.ctrlKey && !e.metaKey : e.metaKey && !e.ctrlKey) : (e.ctrlKey && !e.metaKey);
                if (k.o.win && mac) return false;
                return !!k.c === !!(k.c ? c : (e.ctrlKey || e.metaKey)) && !!k.a === e.altKey && !!k.s === e.shiftKey;
            });
            if (!entry || e.defaultPrevented && !entry.o.force) return;
            if (Common.Utils.ModalWindow.isVisible() && !entry.o.anywhere || $('#file-menu-panel').is(':visible') && !/^F12|KeyO$/.test(e.code)) return;
            var t = e.target, inGrid = !t || t === document.body || $(t).closest('#editor_sdk').length > 0;
            if (!inGrid && !entry.o.anywhere && !(entry.o.inEdit && t && t.id === 'ce-cell-content')) return;
            if (this.api.isCellEdited && !entry.o.inEdit) return;
            var handled = false;
            try { handled = entry.fn.call(this, e) !== true; } catch (err) { console.warn('ExcelShortcuts', err); }
            if (handled) {
                e.preventDefault();
                e.stopImmediatePropagation();
            }
        },

        // ---- commands ---------------------------------------------------------------------
        backstage: function (panel) {
            var lm = this.ctrl('LeftMenu');
            if (!lm || !lm.leftMenu) return true;
            Common.UI.Menu.Manager.hideAll();
            lm.leftMenu.showMenu('file');
            var fm = lm.leftMenu.menuFile;
            panel && fm && fm.panels && fm.panels[panel] && fm.selectMenu(panel);
        },

        toggleFormulaBar: function () {
            var $b = $('#ce-btn-expand');
            $b.length && $b[0].click();
        },

        cycleRegions: function (dir) {
            // grid -> ribbon -> sheet tabs -> grid, like Excel
            var regions = [
                function () { return $(); },
                function () { return $('#toolbar .ribtab.active > a'); },
                function () { return $('#statusbar li.active, #status-btn-addtab').filter(':visible').first(); }
            ];
            var a = document.activeElement, cur = 0;
            if (a && $(a).closest('#toolbar').length) cur = 1;
            else if (a && $(a).closest('#statusbar').length) cur = 2;
            var next = (cur + dir + regions.length) % regions.length, $el = regions[next]();
            if (next === 0 || !$el.length) {
                this.api.asc_enableKeyEvents(true);
                a && a.blur && a.blur();
                Common.NotificationCenter.trigger('edit:complete');
                return;
            }
            this.api.asc_enableKeyEvents(false);
            $el.attr('tabindex') === undefined && $el.attr('tabindex', '-1');
            $el[0].focus();
        },

        zoom: function (d) {
            var z = Math.round((this.api.asc_getZoom() + d) * 10) / 10;
            this.api.asc_setZoom(Math.max(0.1, Math.min(4, z)));
        },

        goTo: function () {
            // Excel's Go To: the reference goes in the Name Box
            var $nb = $('#ce-cell-name');
            if (!$nb.length || $nb.is('[disabled]')) return true;
            this.api.asc_enableKeyEvents(false);
            $nb.focus().select();
        },

        formulaBar: function () {
            var $c = $('#ce-cell-content');
            if (!$c.length || !this.mode.isEdit) return true;
            $c.focus();
            var el = $c[0];
            el.setSelectionRange && el.setSelectionRange(el.value.length, el.value.length);
        },

        fill: function (type) {
            if (!this.tb || (this.tb.btnFillNumbers && this.tb.btnFillNumbers.isDisabled())) return;
            this.api.asc_FillCells(type);
            Common.NotificationCenter.trigger('edit:complete', this.tb);
        },

        hide: function (what, hide) {
            var info = this.api.asc_getCellInfo();
            if (info && info.asc_getLocked && info.asc_getLocked()) return;
            what === 'rows' ? (hide ? this.api.asc_hideRows() : this.api.asc_showRows())
                            : (hide ? this.api.asc_hideColumns() : this.api.asc_showColumns());
            Common.NotificationCenter.trigger('edit:complete', this.tb);
        },

        group: function (on) {
            var info = this.api.asc_getCellInfo(), T = Asc.c_oAscSelectionType,
                type = info && info.asc_getSelectionType(), rows = type !== T.RangeCol;
            on ? this.api.asc_group(rows) : this.api.asc_ungroup(rows);
            Common.NotificationCenter.trigger('edit:complete', this.tb);
        },

        newSheet: function () {
            var sbc = this.ctrl('Statusbar');
            if (!sbc || !sbc.createSheetName) return true;
            var me = this, name = sbc.createSheetName([]);
            setTimeout(function () { me.api.asc_insertWorksheet([name]); }, 1);
        },

        table_: function () {
            // Excel's Ctrl+T, like the editor's Ctrl+L: format the region as a table
            var tbc = this.tbc, s = tbc && tbc._state;
            if (!tbc || !tbc.editMode || !s || s.multiselect || s.wsLock || !tbc.appConfig.canModifyFilter) return;
            var info = this.api.asc_getCellInfo(), fi = info.asc_getAutoFilterInfo(), ti = info.asc_getFormatTableInfo();
            fi = fi ? fi.asc_getIsAutoFilter() : null;
            fi !== null && !ti && tbc._setTableFormat(this.api.asc_getDefaultTableStyle());
        },

        spelling: function () {
            var lm = this.ctrl('LeftMenu'), v = lm && lm.leftMenu, b = v && v.btnSpellcheck;
            if (!b || b.isDisabled() || !b.isVisible()) return true;
            !b.pressed && b.toggle(true);
            v.onBtnMenuClick(b);
            Common.NotificationCenter.trigger('layout:changed', 'leftmenu');
        },

        comment: function () {
            var dh = this.ctrl('DocumentHolder');
            if (!dh || !dh.onAddComment || this.mode.canComments === false) return true;
            dh.onAddComment();
        },

        deleteCells: function () {
            var tb = this.tb;
            if (!tb || !tb.btnDeleteCell || tb.btnDeleteCell.isDisabled()) return;
            var me = this, info = this.api.asc_getCellInfo(), T = Asc.c_oAscSelectionType, type = info.asc_getSelectionType();
            if (type === T.RangeRow || type === T.RangeCol) {
                this.api.asc_deleteCells(type === T.RangeRow ? Asc.c_oAscDeleteOptions.DeleteRows : Asc.c_oAscDeleteOptions.DeleteColumns);
                Common.NotificationCenter.trigger('edit:complete', tb);
                return;
            }
            var items = tb.btnDeleteCell.menu.getItems(true), arr = [];
            for (var i = 0; i < 4 && i < items.length; i++)
                arr.push({caption: items[i].caption, value: items[i].value, disabled: items[i].isDisabled()});
            new Common.Views.OptionsDialog({
                title: this.tbc.txtDeleteCells,
                items: arr,
                handler: function (dlg, result) {
                    result === 'ok' && me.api.asc_deleteCells(dlg.getSettings());
                    Common.NotificationCenter.trigger('edit:complete', tb);
                }
            }).show();
        },

        click: function (name, toggle) {
            var b = this.tb && this.tb[name];
            if (!b || b.isDisabled()) return;
            if (toggle && b.options.enableToggle) b.toggle(!b.pressed, true);
            b.trigger('click', b, {});
        },

        borders: function (id) {
            var b = this.tb && this.tb.btnBorders;
            if (!b || b.isDisabled() || !b.menu || !b.menu.items) return;
            var item = _.find(b.menu.items, function (it) { return it.options && it.options.borderId === id; });
            item && this.tbc.onBordersMenu(b.menu, item);
        },

        calculate: function () {
            this.api.asc_calculate(Asc.c_oAscCalculateType.All);
        },

        macros: function () {
            var vt = this.ctrl('ViewTab');
            if (!vt || !vt.onClickMacros || !this.mode.isEdit) return true;
            vt.onClickMacros();
        },

        find: function (replace) {
            var lmc = this.ctrl('LeftMenu');
            if (!lmc) return true;
            lmc.onShowHideSearch(true);
            if (replace) {
                var panel = lmc.leftMenu && lmc.leftMenu.panelSearch;
                setTimeout(function () { panel && panel.focus && panel.focus('replace'); }, 100);
            }
        },

        filter: function () {
            var b = this.tb && this.tb.btnsSetAutofilter && this.tb.btnsSetAutofilter[0];
            if (!b || b.isDisabled()) return;
            b.toggle(!b.pressed, true);
            b.trigger('click', b, {});
        },

        customSort: function () {
            var b = this.tb && this.tb.btnCustomSort;
            b = _.isArray(b) ? b[0] : b;
            b && !b.isDisabled() && b.trigger('click', b, {});
        },

        autoSum: function () {
            var b = this.tb && this.tb.btnInsertFormula;
            b && !b.isDisabled() && b.trigger('click', b, {});
        },

        // Excel's Ctrl+Shift+Space: like Ctrl+A (the region, then the sheet), all the objects when one is selected
        selectAll: function () {
            var wb = this.api.wb, ws = wb && wb.getWorksheet();
            if (!ws) return true;
            var sel = ws.objectRender && ws.objectRender.getSelectedGraphicObjects();
            if (sel && sel.length) {
                ws.objectRender.controller.selectAll();
                ws.objectRender.controller.drawingObjects.sendGraphicObjectProps();
            } else wb.controller.handlers.trigger('selectAllByRange');
        },

        names: function (value) {
            if (!this.tbc || !this.tbc.onNamedRangeMenu) return true;
            this.tbc.onNamedRangeMenu(null, {value: value, options: {}});
        },

        chart: function () {
            // Excel's Alt+F1: a clustered column chart of the data
            if (!this.tbc || !this.tbc.onSelectChart || (this.tb.btnInsertChart && this.tb.btnInsertChart.isDisabled())) return;
            this.tbc.onSelectChart('bar', Asc.c_oAscChartTypeSettings.barNormal);
        },

        // Paste Special: Excel's access keys for the options (A, F, V, T, C, X, W, R, U)
        pasteSpecialKey: function (e) {
            if (e.ctrlKey || e.metaKey || e.altKey || e.key.length !== 1) return;
            var map = {a: 'paste-radio-all', f: 'paste-radio-formulas', v: 'paste-radio-values', t: 'paste-radio-formats',
                       c: 'paste-radio-comments', x: 'paste-radio-without-borders', w: 'paste-radio-col-width',
                       r: 'paste-radio-formula-num-format', u: 'paste-radio-value-num-format'};
            var id = map[e.key.toLowerCase()], el = id && document.getElementById(id);
            if (!el || !el.getClientRects().length) return;
            // not while typing in a text field (the grid's own input does not count)
            if (e.target && e.target.id !== 'area_id' && /^(INPUT|TEXTAREA)$/.test(e.target.tagName) && e.target.type !== 'radio') return;
            var input = el.querySelector('input');
            if (!input) return;
            input.click();
            e.preventDefault();
            e.stopImmediatePropagation();
            return true;
        },

        toggleRibbon: function () {
            var rt = this.ctrl('RibbonTabs');
            rt && rt.toggleRibbon && rt.toggleRibbon();
        }
    });
});

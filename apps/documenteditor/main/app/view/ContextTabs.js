/*
 * SPDX-FileCopyrightText: 2026 ElixirOffice contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 *
 * DE.Views.ContextTabs: Word-like contextual ribbon tabs of the document editor.
 *  - Table Design and Layout (table) when the cursor is in a table;
 *  - Picture Format when a picture is selected;
 *  - Shape Format when a shape or a text box is selected (or its text edited);
 *  - Equation when the cursor is in an equation.
 * The Header & Footer tab already exists (HeaderFooterTab); the controller only
 * restyles it as a contextual tab.
 *
 * Commands that already exist as ribbon buttons elsewhere (Wrap Text, Align,
 * Bring Forward...) are not re-implemented: see Common.UI.RibbonMirror.
 */

define([
    'core',
    'common/main/lib/component/Button',
    'common/main/lib/component/CheckBox',
    'common/main/lib/component/MetricSpinner',
    'common/main/lib/component/ComboDataView',
    'common/main/lib/component/ColorButton',
    'common/main/lib/component/RibbonMirror'
], function () {
    'use strict';

    DE.Views.ContextTabs = Common.UI.BaseView.extend(_.extend((function () {

        function group(label, launcher, inner, cls) {
            return '<div class="group' + (cls ? ' ' + cls : '') + '" data-group="' + label + '"' + (launcher ? ' data-launcher="' + launcher + '"' : '') + '>' + inner + '</div>';
        }
        var SEP = '<div class="separator long"></div>',
            big = function (id) { return '<span class="btn-slot text x-huge" id="' + id + '"></span>'; },
            small = function (id, cls) { return '<span class="btn-slot' + (cls ? ' ' + cls : '') + '" id="' + id + '"></span>'; },
            elset = function (inner, style) { return '<div class="elset"' + (style ? ' style="' + style + '"' : '') + '>' + inner + '</div>'; },
            column = function (inner) { return '<div class="elset-column">' + inner + '</div>'; },
            spinRow = function (lbl, spin) {
                return '<div class="elset" style="display: flex; align-items: baseline;">' +
                    '<span class="btn-slot text font-size-normal margin-right-6" id="' + lbl + '" style="flex-grow:1;"></span>' +
                    '<span class="btn-slot text spinner" id="' + spin + '"></span></div>';
            },
            panel = function (tab, inner) {
                return '<section class="panel elx-context-panel" data-tab="' + tab + '" role="tabpanel" aria-labelledby="' + tab + '">' + inner + '</section>';
            };

        var templates = {
            tabledesign: panel('tabledesign',
                group('tablestyleoptions', null,
                    column(elset(small('elx-chk-tbl-header', 'text')) + elset(small('elx-chk-tbl-total', 'text')) + elset(small('elx-chk-tbl-banded', 'text'))) +
                    column(elset(small('elx-chk-tbl-first', 'text')) + elset(small('elx-chk-tbl-last', 'text')) + elset(small('elx-chk-tbl-colbanded', 'text'))),
                    'small elx-table-options') + SEP +
                '<div class="group small flex elx-tbl-styles" id="elx-tbl-styles" data-group="tablestyles" style="width: 100%; min-width: 160px;" data-group-width="100%"></div>' +
                group('tablestyles', null, big('elx-btn-tbl-shading')) + SEP +
                group('borders', 'borders',
                    column(elset(small('elx-cmb-tbl-border-size')) + elset(small('elx-btn-tbl-pen-color', 'text'))) +
                    big('elx-btn-tbl-borders'), 'small elx-borders')
            ),
            tablelayout: panel('tablelayout',
                group('table', null, big('elx-btn-tbl-select') + big('elx-btn-tbl-gridlines') + big('elx-btn-tbl-props')) + SEP +
                group('draw', null, big('elx-btn-tbl-draw') + big('elx-btn-tbl-eraser')) + SEP +
                group('rowscolumns', 'cells', big('elx-btn-tbl-delete') + big('elx-btn-tbl-ins-above') + big('elx-btn-tbl-ins-below') +
                    big('elx-btn-tbl-ins-left') + big('elx-btn-tbl-ins-right')) + SEP +
                group('merge', null, big('elx-btn-tbl-merge') + big('elx-btn-tbl-split')) + SEP +
                group('cellsize', 'cellsize', big('elx-btn-tbl-autofit')) +
                group('cellsize', null, spinRow('elx-lbl-tbl-height', 'elx-spin-tbl-height') + spinRow('elx-lbl-tbl-width', 'elx-spin-tbl-width'), 'small') +
                group('cellsize', null, elset(small('elx-btn-tbl-dist-rows', 'text')) + elset(small('elx-btn-tbl-dist-cols', 'text')), 'small') + SEP +
                group('alignment', null,
                    '<div class="elx-align-grid">' +
                        elset(small('elx-btn-tbl-al-0') + small('elx-btn-tbl-al-1') + small('elx-btn-tbl-al-2')) +
                        elset(small('elx-btn-tbl-al-3') + small('elx-btn-tbl-al-4') + small('elx-btn-tbl-al-5')) +
                        elset(small('elx-btn-tbl-al-6') + small('elx-btn-tbl-al-7') + small('elx-btn-tbl-al-8')) +
                    '</div>', 'small') +
                group('alignment', 'cellmargins', big('elx-btn-tbl-direction') + big('elx-btn-tbl-margins')) + SEP +
                group('data', null, big('elx-btn-tbl-repeat-header') + big('elx-btn-tbl-to-text') + big('elx-btn-tbl-formula'))
            ),
            pictureformat: panel('pictureformat',
                group('adjust', null, big('elx-btn-pic-transparency') + big('elx-btn-pic-change') + big('elx-btn-pic-reset')) + SEP +
                group('picturestyles', 'picturestyles', big('elx-btn-pic-border') + big('elx-btn-pic-effects')) + SEP +
                group('accessibility', null, big('elx-btn-pic-alttext')) + SEP +
                group('arrange', null, big('elx-btn-pic-position') + big('elx-btn-pic-wrap') + big('elx-btn-pic-forward') +
                    big('elx-btn-pic-backward') + big('elx-btn-pic-align') + big('elx-btn-pic-group') + big('elx-btn-pic-rotate')) + SEP +
                group('size', 'picsize', big('elx-btn-pic-crop')) +
                group('size', null, spinRow('elx-lbl-pic-height', 'elx-spin-pic-height') + spinRow('elx-lbl-pic-width', 'elx-spin-pic-width'), 'small')
            ),
            shapeformat: panel('shapeformat',
                group('insertshapes', null, big('elx-btn-shp-shapes') + big('elx-btn-shp-edit') + big('elx-btn-shp-textbox')) + SEP +
                group('shapestyles', 'shapestyles', big('elx-btn-shp-fill') + big('elx-btn-shp-outline') + big('elx-btn-shp-effects')) + SEP +
                group('text', null, big('elx-btn-shp-direction') + big('elx-btn-shp-aligntext')) + SEP +
                group('accessibility', null, big('elx-btn-shp-alttext')) + SEP +
                group('arrange', null, big('elx-btn-shp-position') + big('elx-btn-shp-wrap') + big('elx-btn-shp-forward') +
                    big('elx-btn-shp-backward') + big('elx-btn-shp-align') + big('elx-btn-shp-group') + big('elx-btn-shp-rotate')) + SEP +
                group('size', 'shpsize', spinRow('elx-lbl-shp-height', 'elx-spin-shp-height') + spinRow('elx-lbl-shp-width', 'elx-spin-shp-width'), 'small')
            ),
            equation: panel('equation',
                group('tools', null, big('elx-btn-eq-new') + big('elx-btn-eq-convert') + big('elx-btn-eq-display')) + SEP +
                group('structures', null, '<span class="elx-eq-structures"></span>')
            )
        };

        // Word's position presets: [horizontal align, vertical align]
        var POSITIONS = [
            ['Left', 'Top'], ['Center', 'Top'], ['Right', 'Top'],
            ['Left', 'Center'], ['Center', 'Center'], ['Right', 'Center'],
            ['Left', 'Bottom'], ['Center', 'Bottom'], ['Right', 'Bottom']
        ];

        return {
            options: {},

            initialize: function (options) {
                Common.UI.BaseView.prototype.initialize.call(this);
                this.toolbar = options.toolbar;
                this.mode = options.mode;
                this.controls = {tabledesign: [], tablelayout: [], pictureformat: [], shapeformat: [], equation: []};
                this.lockedControls = [];
                this.spinners = [];
                this.panels = {};
            },

            _baseLock: function () {
                var _set = Common.enumLock;
                return [_set.previewReviewMode, _set.viewFormMode, _set.lostConnect, _set.disableOnStart, _set.docLockView, _set.docLockForms, _set.docLockComments, _set.viewMode, 'elx-ctx-locked'];
            },

            _btn: function (tab, slot, cfg) {
                var btn = new Common.UI.Button(_.extend({
                    cls: 'btn-toolbar x-huge icon-top',
                    lock: this._baseLock(),
                    dataHint: '1',
                    dataHintDirection: 'bottom',
                    dataHintOffset: 'small'
                }, cfg));
                this._register(tab, slot, btn);
                return btn;
            },

            _small: function (tab, slot, cfg) {
                return this._btn(tab, slot, _.extend({cls: 'btn-toolbar', dataHintOffset: '0, -6'}, cfg));
            },

            _register: function (tab, slot, cmp) {
                var $slot = this.panels[tab].find('#' + slot);
                !cmp.rendered && cmp.render($slot);
                this.controls[tab].push(cmp);
                this.lockedControls.push(cmp);
                return cmp;
            },

            _mirror: function (tab, slot, src, opts) {
                var twin = Common.UI.RibbonMirror(src, opts);
                twin.options.lock = this._baseLock();
                this._register(tab, slot, twin);
                return twin;
            },

            _spinner: function (tab, lblSlot, slot, caption) {
                var lbl = new Common.UI.Label({caption: caption, lock: this._baseLock()});
                this._register(tab, lblSlot, lbl);
                var spin = new Common.UI.MetricSpinner({
                    el: this.panels[tab].find('#' + slot),
                    step: .1, width: 85, defaultUnit: 'cm', value: '1 cm', maxValue: 55.88, minValue: 0,
                    lock: this._baseLock(), dataHint: '1', dataHintDirection: 'bottom', dataHintOffset: 'big'
                });
                this._register(tab, slot, spin);
                spin.lbl = lbl;
                this.spinners.push(spin);
                return spin;
            },

            _check: function (tab, slot, caption) {
                var ch = new Common.UI.CheckBox({labelText: caption, lock: this._baseLock(), dataHint: '1', dataHintDirection: 'left', dataHintOffset: 'small'});
                this._register(tab, slot, ch);
                return ch;
            },

            createPanels: function () {
                var me = this;
                _.each(templates, function (tpl, tab) { me.panels[tab] = $(tpl); });
                this.createTableDesign();
                this.createTableLayout();
                this.createPictureFormat();
                this.createShapeFormat();
                this.createEquation();
                return this.panels;
            },

            // ---- Table Design ---------------------------------------------------------
            createTableDesign: function () {
                var me = this, t = 'tabledesign';
                this.chHeader = this._check(t, 'elx-chk-tbl-header', this.txtHeaderRow);
                this.chTotal = this._check(t, 'elx-chk-tbl-total', this.txtTotalRow);
                this.chBanded = this._check(t, 'elx-chk-tbl-banded', this.txtBandedRows);
                this.chFirst = this._check(t, 'elx-chk-tbl-first', this.txtFirstColumn);
                this.chLast = this._check(t, 'elx-chk-tbl-last', this.txtLastColumn);
                this.chColBanded = this._check(t, 'elx-chk-tbl-colbanded', this.txtBandedColumns);

                var itemWidth = 66, itemHeight = 48;
                this.tableStyles = new Common.UI.ComboDataView({
                    cls: 'combo-styles elx-combo-table-styles',
                    itemWidth: itemWidth,
                    itemHeight: itemHeight,
                    style: 'min-width: 150px; max-width: 520px;',
                    lock: this._baseLock(),
                    enableKeyEvents: true,
                    autoWidth: true,
                    delayRenderTips: true,
                    groups: new Common.UI.DataViewGroupStore(),
                    menuMaxHeight: 400,
                    itemTemplate: _.template('<div class="style" id="<%= id %>"><img src="<%= imageUrl %>" width="' + itemWidth + '" height="' + itemHeight + '"></div>'),
                    beforeOpenHandler: function () {
                        var cmp = this, menu = cmp.openButton.menu, columnCount = 7;
                        if (menu.cmpEl) {
                            var itemEl = $(cmp.cmpEl.find('.dataview.inner .style').get(0)).parent();
                            var itemMargin = 8,
                                w = itemEl.is(':visible') ? parseFloat(itemEl.css('width')) : (cmp.itemWidth + 6);
                            menu.menuAlignEl = cmp.cmpEl;
                            menu.menuAlign = 'tl-tl';
                            var menuWidth = columnCount * (itemMargin + w) + 17;
                            if (menuWidth > Common.Utils.innerWidth())
                                menuWidth = Math.max(Math.floor((Common.Utils.innerWidth() - 17) / (itemMargin + w)), 2) * (itemMargin + w) + 17;
                            var offset = cmp.cmpEl.width() - cmp.openButton.$el.width() - Math.min(menuWidth, Common.Utils.getOffset(cmp.openButton.$el).left) - 1;
                            menu.setOffset(Common.UI.isRTL() ? cmp.openButton.$el.width() : Math.min(offset, 0));
                            menu.cmpEl.css({'width': menuWidth, 'min-height': cmp.cmpEl.height()});
                        }
                        cmp.menuPicker.scroller && cmp.menuPicker.scroller.update({includePadding: true, suppressScrollX: true});
                        cmp.removeTips && cmp.removeTips();
                    },
                    dataHint: '1',
                    dataHintDirection: 'bottom',
                    dataHintOffset: '-16, 0'
                });
                this._register(t, 'elx-tbl-styles', this.tableStyles);

                this.btnShading = new Common.UI.ButtonColored({
                    cls: 'btn-toolbar x-huge icon-top',
                    iconCls: 'toolbar__icon btn-elx-shading',
                    caption: this.txtShading,
                    split: true,
                    menu: true,
                    transparent: true,
                    eyeDropper: true,
                    colorLine: false,
                    lock: this._baseLock(),
                    dataHint: '1', dataHintDirection: 'bottom', dataHintOffset: 'small'
                });
                this._register(t, 'elx-btn-tbl-shading', this.btnShading);

                this.cmbBorderSize = new Common.UI.ComboBorderSize({
                    style: 'width: 93px;',
                    lock: this._baseLock(),
                    dataHint: '1', dataHintDirection: 'bottom', dataHintOffset: 'big',
                    ariaLabel: this.txtPenWeight
                });
                this._register(t, 'elx-cmb-tbl-border-size', this.cmbBorderSize);
                this.BorderSize = this.cmbBorderSize.store.at(1).get('value');
                this.cmbBorderSize.setValue(this.BorderSize);

                this.btnPenColor = new Common.UI.ButtonColored({
                    cls: 'btn-toolbar',
                    iconCls: 'toolbar__icon btn-elx-pen-color',
                    caption: this.txtPenColor,
                    menu: true,
                    auto: true,
                    color: 'auto',
                    eyeDropper: true,
                    lock: this._baseLock(),
                    dataHint: '1', dataHintDirection: 'bottom', dataHintOffset: '0, -6'
                });
                this._register(t, 'elx-btn-tbl-pen-color', this.btnPenColor);

                var bi = function (value, cls, caption) { return {caption: caption, iconCls: 'menu__icon ' + cls, value: value}; };
                this.btnBorders = this._btn(t, 'elx-btn-tbl-borders', {
                    iconCls: 'toolbar__icon btn-elx-borders',
                    caption: this.txtBorders,
                    split: true,
                    menu: new Common.UI.Menu({
                        cls: 'shifted-right',
                        items: [
                            bi('b', 'btn-border-bottom', this.txtBorderBottom),
                            bi('t', 'btn-border-top', this.txtBorderTop),
                            bi('l', 'btn-border-left', this.txtBorderLeft),
                            bi('r', 'btn-border-right', this.txtBorderRight),
                            {caption: '--'},
                            bi('', 'btn-border-no', this.txtBorderNone),
                            bi('lrtbcm', 'btn-border-all', this.txtBorderAll),
                            bi('lrtb', 'btn-border-out', this.txtBorderOutside),
                            bi('cm', 'btn-border-inside', this.txtBorderInside),
                            bi('m', 'btn-border-insidehor', this.txtBorderInsideH),
                            bi('c', 'btn-border-insidevert', this.txtBorderInsideV),
                            {caption: '--'},
                            {caption: this.txtBordersShading, value: 'advanced'}
                        ]
                    })
                });
                this.btnBorders.options.lastBorder = 'b';
            },

            // ---- Table Layout ---------------------------------------------------------
            createTableLayout: function () {
                var me = this, t = 'tablelayout';
                this.btnTblSelect = this._btn(t, 'elx-btn-tbl-select', {
                    iconCls: 'toolbar__icon btn-elx-tbl-select', caption: this.txtSelect,
                    menu: new Common.UI.Menu({items: [
                        {caption: this.txtSelectCell, value: 'cell'},
                        {caption: this.txtSelectColumn, value: 'column'},
                        {caption: this.txtSelectRow, value: 'row'},
                        {caption: this.txtSelectTable, value: 'table'}
                    ]})
                });
                this.btnTblGridlines = this._btn(t, 'elx-btn-tbl-gridlines', {iconCls: 'toolbar__icon btn-elx-tbl-gridlines', caption: this.txtViewGridlines, enableToggle: true});
                this.btnTblProps = this._btn(t, 'elx-btn-tbl-props', {iconCls: 'toolbar__icon btn-properties', caption: this.txtProperties});
                this.btnTblDraw = this._btn(t, 'elx-btn-tbl-draw', {iconCls: 'toolbar__icon btn-elx-tbl-draw', caption: this.txtDrawTable, enableToggle: true});
                this.btnTblEraser = this._btn(t, 'elx-btn-tbl-eraser', {iconCls: 'toolbar__icon btn-eraser', caption: this.txtEraser, enableToggle: true});
                this.btnTblDelete = this._btn(t, 'elx-btn-tbl-delete', {
                    iconCls: 'toolbar__icon btn-elx-tbl-delete', caption: this.txtDelete,
                    menu: new Common.UI.Menu({items: [
                        {caption: this.txtDeleteCells, value: 'cells'},
                        {caption: this.txtDeleteColumns, value: 'column'},
                        {caption: this.txtDeleteRows, value: 'row'},
                        {caption: this.txtDeleteTable, value: 'table'}
                    ]})
                });
                this.btnTblInsAbove = this._btn(t, 'elx-btn-tbl-ins-above', {iconCls: 'toolbar__icon btn-elx-tbl-ins-above', caption: this.txtInsertAbove});
                this.btnTblInsBelow = this._btn(t, 'elx-btn-tbl-ins-below', {iconCls: 'toolbar__icon btn-elx-tbl-ins-below', caption: this.txtInsertBelow});
                this.btnTblInsLeft = this._btn(t, 'elx-btn-tbl-ins-left', {iconCls: 'toolbar__icon btn-elx-tbl-ins-left', caption: this.txtInsertLeft});
                this.btnTblInsRight = this._btn(t, 'elx-btn-tbl-ins-right', {iconCls: 'toolbar__icon btn-elx-tbl-ins-right', caption: this.txtInsertRight});
                this.btnTblMerge = this._btn(t, 'elx-btn-tbl-merge', {iconCls: 'toolbar__icon btn-elx-tbl-merge', caption: this.txtMergeCells});
                this.btnTblSplit = this._btn(t, 'elx-btn-tbl-split', {iconCls: 'toolbar__icon btn-elx-tbl-split', caption: this.txtSplitCells});
                this.btnTblAutofit = this._btn(t, 'elx-btn-tbl-autofit', {
                    iconCls: 'toolbar__icon btn-elx-tbl-autofit', caption: this.txtAutoFit,
                    menu: new Common.UI.Menu({items: [
                        {caption: this.txtAutoFitContents, value: 'contents'},
                        {caption: this.txtAutoFitWindow, value: 'window'},
                        {caption: this.txtFixedColumnWidth, value: 'fixed'}
                    ]})
                });
                this.spnTblHeight = this._spinner(t, 'elx-lbl-tbl-height', 'elx-spin-tbl-height', this.txtHeight);
                this.spnTblWidth = this._spinner(t, 'elx-lbl-tbl-width', 'elx-spin-tbl-width', this.txtWidth);
                this.btnTblDistRows = this._small(t, 'elx-btn-tbl-dist-rows', {iconCls: 'toolbar__icon btn-distribute-rows', caption: this.txtDistributeRows});
                this.btnTblDistCols = this._small(t, 'elx-btn-tbl-dist-cols', {iconCls: 'toolbar__icon btn-distribute-columns', caption: this.txtDistributeColumns});

                // 3x3 cell alignment grid, like Word: [paragraph alignment, cell vertical alignment]
                var names = ['TopLeft', 'TopCenter', 'TopRight', 'CenterLeft', 'Center', 'CenterRight', 'BottomLeft', 'BottomCenter', 'BottomRight'],
                    icons = ['top-left', 'top-center', 'top-right', 'middle-left', 'middle-center', 'middle-right', 'bottom-left', 'bottom-center', 'bottom-right'];
                this.btnsCellAlign = [];
                for (var i = 0; i < 9; i++) {
                    var b = this._small(t, 'elx-btn-tbl-al-' + i, {
                        iconCls: 'toolbar__icon btn-elx-cell-' + icons[i],
                        enableToggle: true, allowDepress: false, toggleGroup: 'elx-cell-align',
                        hint: this['txtAlign' + names[i]]
                    });
                    b.options.halign = [1, 2, 0][i % 3]; // AscCommon.align_Left=1, Center=2, Right=0
                    b.options.valign = [Asc.c_oAscVertAlignJc.Top, Asc.c_oAscVertAlignJc.Center, Asc.c_oAscVertAlignJc.Bottom][Math.floor(i / 3)];
                    this.btnsCellAlign.push(b);
                }
                this.btnTblDirection = this._btn(t, 'elx-btn-tbl-direction', {
                    iconCls: 'toolbar__icon btn-elx-text-direction', caption: this.txtTextDirection,
                    menu: new Common.UI.Menu({cls: 'ppm-toolbar shifted-right', items: [
                        {caption: this.txtDirHorizontal, iconCls: 'menu__icon btn-text-orient-hor', checkable: true, checkmark: false, toggleGroup: 'elx-tbl-dir', direction: Asc.c_oAscCellTextDirection.LRTB},
                        {caption: this.txtDirRotateDown, iconCls: 'menu__icon btn-text-orient-rdown', checkable: true, checkmark: false, toggleGroup: 'elx-tbl-dir', direction: Asc.c_oAscCellTextDirection.TBRL},
                        {caption: this.txtDirRotateUp, iconCls: 'menu__icon btn-text-orient-rup', checkable: true, checkmark: false, toggleGroup: 'elx-tbl-dir', direction: Asc.c_oAscCellTextDirection.BTLR}
                    ]})
                });
                this.btnTblMargins = this._btn(t, 'elx-btn-tbl-margins', {iconCls: 'toolbar__icon btn-elx-cell-margins', caption: this.txtCellMargins});
                this.btnTblRepeatHeader = this._btn(t, 'elx-btn-tbl-repeat-header', {iconCls: 'toolbar__icon btn-elx-tbl-repeat-header', caption: this.txtRepeatHeader, enableToggle: true});
                this.btnTblToText = this._btn(t, 'elx-btn-tbl-to-text', {iconCls: 'toolbar__icon btn-elx-tbl-to-text', caption: this.txtConvertToText});
                this.btnTblFormula = this._btn(t, 'elx-btn-tbl-formula', {iconCls: 'toolbar__icon btn-elx-formula', caption: this.txtFormula});
            },

            // ---- Picture Format and Shape Format (shared Arrange group) ---------------
            _createArrange: function (t, prefix) {
                var tb = this.toolbar, res = {};
                res.position = this._btn(t, 'elx-btn-' + prefix + '-position', {
                    iconCls: 'toolbar__icon btn-elx-position', caption: this.txtPosition,
                    menu: new Common.UI.Menu({cls: 'elx-position-menu', items: this._positionItems()})
                });
                res.wrap = this._mirror(t, 'elx-btn-' + prefix + '-wrap', function () { return tb.btnImgWrapping; }, {caption: this.txtWrapText});
                res.forward = this._mirror(t, 'elx-btn-' + prefix + '-forward', function () { return tb.btnImgForward; });
                res.backward = this._mirror(t, 'elx-btn-' + prefix + '-backward', function () { return tb.btnImgBackward; });
                res.align = this._mirror(t, 'elx-btn-' + prefix + '-align', function () { return tb.btnImgAlign; });
                res.group = this._mirror(t, 'elx-btn-' + prefix + '-group', function () { return tb.btnImgGroup; });
                res.rotate = this._btn(t, 'elx-btn-' + prefix + '-rotate', {
                    iconCls: 'toolbar__icon btn-elx-rotate', caption: this.txtRotate,
                    menu: new Common.UI.Menu({items: [
                        {caption: this.txtRotateRight, iconCls: 'menu__icon btn-rotate-90', value: 'r90'},
                        {caption: this.txtRotateLeft, iconCls: 'menu__icon btn-rotate-270', value: 'r270'},
                        {caption: this.txtFlipVertical, iconCls: 'menu__icon btn-flip-vert', value: 'fv'},
                        {caption: this.txtFlipHorizontal, iconCls: 'menu__icon btn-flip-hor', value: 'fh'},
                        {caption: '--'},
                        {caption: this.txtMoreRotation, value: 'more'}
                    ]})
                });
                return res;
            },

            _positionItems: function () {
                var me = this, items = [{caption: this.txtInLine, value: 'inline', iconCls: 'menu__icon btn-small-wrap-inline'}, {caption: '--'}];
                _.each(POSITIONS, function (p, i) {
                    items.push({caption: me['txtPos' + p[1] + p[0]], value: i, iconCls: 'menu__icon btn-elx-pos-' + i});
                });
                items.push({caption: '--'});
                items.push({caption: this.txtMoreLayout, value: 'more'});
                return items;
            },

            _createOutline: function (t, slot, caption, iconCls) {
                var weight = new Common.UI.MenuItem({caption: this.txtWeight, menu: new Common.UI.Menu({menuAlign: 'tl-tr', items: _.map([0.25, 0.5, 0.75, 1, 1.5, 2.25, 3, 4.5, 6], function (v) {
                    return {caption: v + ' pt', value: v, checkable: true, toggleGroup: 'elx-' + slot + '-weight'};
                })})});
                var btn = new Common.UI.ButtonColored({
                    cls: 'btn-toolbar x-huge icon-top', iconCls: iconCls, caption: caption,
                    split: true, menu: true, transparent: true, eyeDropper: true, colorLine: false, lock: this._baseLock(),
                    additionalItemsAfter: [{caption: '--'}, weight],
                    dataHint: '1', dataHintDirection: 'bottom', dataHintOffset: 'small'
                });
                btn.mnuWeight = weight;
                this._register(t, slot, btn);
                return btn;
            },

            _createEffects: function (t, slot, caption) {
                return this._btn(t, slot, {
                    iconCls: 'toolbar__icon btn-elx-shape-effects', caption: caption,
                    menu: new Common.UI.Menu({items: [
                        {caption: this.txtNoShadow, value: 'none', checkable: true, toggleGroup: 'elx-' + slot},
                        {caption: this.txtShadowOuter, value: 'outerShdw', checkable: true, toggleGroup: 'elx-' + slot},
                        {caption: '--'},
                        {caption: this.txtShadowOptions, value: 'options'}
                    ]})
                });
            },

            createPictureFormat: function () {
                var t = 'pictureformat';
                this.btnPicTransparency = this._btn(t, 'elx-btn-pic-transparency', {
                    iconCls: 'toolbar__icon btn-elx-transparency', caption: this.txtTransparency,
                    menu: new Common.UI.Menu({items: _.map([0, 15, 30, 50, 65, 80, 95], function (v) {
                        return {caption: v + ' %', value: v, checkable: true, toggleGroup: 'elx-pic-transp'};
                    })})
                });
                this.btnPicChange = this._btn(t, 'elx-btn-pic-change', {
                    iconCls: 'toolbar__icon btn-elx-change-picture', caption: this.txtChangePicture,
                    menu: new Common.UI.Menu({items: [
                        {caption: this.txtFromFile, value: 0},
                        {caption: this.txtFromUrl, value: 1},
                        {caption: this.txtFromStorage, value: 2}
                    ]})
                });
                this.btnPicReset = this._btn(t, 'elx-btn-pic-reset', {
                    iconCls: 'toolbar__icon btn-elx-reset-picture', caption: this.txtResetPicture, split: true,
                    menu: new Common.UI.Menu({items: [
                        {caption: this.txtResetPicture, value: 'reset'},
                        {caption: this.txtResetPictureSize, value: 'size'}
                    ]})
                });
                this.btnPicBorder = this._createOutline(t, 'elx-btn-pic-border', this.txtPictureBorder, 'toolbar__icon btn-elx-shape-outline');
                this.btnPicEffects = this._createEffects(t, 'elx-btn-pic-effects', this.txtPictureEffects);
                this.btnPicAltText = this._btn(t, 'elx-btn-pic-alttext', {iconCls: 'toolbar__icon btn-elx-alt-text', caption: this.txtAltText});
                this.picArrange = this._createArrange(t, 'pic');
                this.btnPicCrop = this._btn(t, 'elx-btn-pic-crop', {
                    iconCls: 'toolbar__icon btn-elx-crop', caption: this.txtCrop, split: true, enableToggle: true,
                    menu: new Common.UI.Menu({items: [
                        {caption: this.txtCrop, value: 0, checkable: true},
                        {caption: this.txtFill, value: 1},
                        {caption: this.txtFit, value: 2}
                    ]})
                });
                this.spnPicHeight = this._spinner(t, 'elx-lbl-pic-height', 'elx-spin-pic-height', this.txtHeight);
                this.spnPicWidth = this._spinner(t, 'elx-lbl-pic-width', 'elx-spin-pic-width', this.txtWidth);
            },

            createShapeFormat: function () {
                var t = 'shapeformat', tb = this.toolbar;
                this.btnShpShapes = this._btn(t, 'elx-btn-shp-shapes', {
                    iconCls: 'toolbar__icon btn-insertshape', caption: tb.capBtnInsShape,
                    menu: new Common.UI.Menu({cls: 'menu-shapes menu-insert-shape', items: []})
                });
                this.btnShpEdit = this._btn(t, 'elx-btn-shp-edit', {
                    iconCls: 'toolbar__icon btn-elx-edit-shape', caption: this.txtEditShape,
                    menu: new Common.UI.Menu({items: [
                        {caption: this.txtChangeShape, value: 'change', menu: new Common.UI.Menu({menuAlign: 'tl-tr', cls: 'menu-shapes', items: []})},
                        {caption: this.txtEditPoints, value: 'points'}
                    ]})
                });
                this.btnShpTextBox = this._mirror(t, 'elx-btn-shp-textbox', function () { return tb.btnInsertText; });
                this.btnShpFill = new Common.UI.ButtonColored({
                    cls: 'btn-toolbar x-huge icon-top', iconCls: 'toolbar__icon btn-elx-shape-fill', caption: this.txtShapeFill,
                    split: true, menu: true, transparent: true, eyeDropper: true, colorLine: false, lock: this._baseLock(),
                    dataHint: '1', dataHintDirection: 'bottom', dataHintOffset: 'small'
                });
                this._register(t, 'elx-btn-shp-fill', this.btnShpFill);
                this.btnShpOutline = this._createOutline(t, 'elx-btn-shp-outline', this.txtShapeOutline, 'toolbar__icon btn-elx-shape-outline');
                this.mnuShpOutlineWeight = this.btnShpOutline.mnuWeight;
                this.btnShpEffects = this._createEffects(t, 'elx-btn-shp-effects', this.txtShapeEffects);
                this.btnShpDirection = this._btn(t, 'elx-btn-shp-direction', {
                    iconCls: 'toolbar__icon btn-elx-text-direction', caption: this.txtTextDirection,
                    menu: new Common.UI.Menu({cls: 'ppm-toolbar shifted-right', items: [
                        {caption: this.txtDirHorizontal, iconCls: 'menu__icon btn-text-orient-hor', checkable: true, checkmark: false, toggleGroup: 'elx-shp-dir', direction: Asc.c_oAscVertDrawingText.normal},
                        {caption: this.txtDirRotateDown, iconCls: 'menu__icon btn-text-orient-rdown', checkable: true, checkmark: false, toggleGroup: 'elx-shp-dir', direction: Asc.c_oAscVertDrawingText.vert},
                        {caption: this.txtDirRotateUp, iconCls: 'menu__icon btn-text-orient-rup', checkable: true, checkmark: false, toggleGroup: 'elx-shp-dir', direction: Asc.c_oAscVertDrawingText.vert270}
                    ]})
                });
                this.btnShpAlignText = this._btn(t, 'elx-btn-shp-aligntext', {
                    iconCls: 'toolbar__icon btn-elx-align-text', caption: this.txtAlignText,
                    menu: new Common.UI.Menu({cls: 'ppm-toolbar shifted-right', items: [
                        {caption: this.txtAlignTop, iconCls: 'menu__icon btn-align-top', checkable: true, checkmark: false, toggleGroup: 'elx-shp-valign', valign: Asc.c_oAscVAlign.Top},
                        {caption: this.txtAlignMiddle, iconCls: 'menu__icon btn-align-middle', checkable: true, checkmark: false, toggleGroup: 'elx-shp-valign', valign: Asc.c_oAscVAlign.Center},
                        {caption: this.txtAlignBottom, iconCls: 'menu__icon btn-align-bottom', checkable: true, checkmark: false, toggleGroup: 'elx-shp-valign', valign: Asc.c_oAscVAlign.Bottom}
                    ]})
                });
                this.btnShpAltText = this._btn(t, 'elx-btn-shp-alttext', {iconCls: 'toolbar__icon btn-elx-alt-text', caption: this.txtAltText});
                this.shpArrange = this._createArrange(t, 'shp');
                this.spnShpHeight = this._spinner(t, 'elx-lbl-shp-height', 'elx-spin-shp-height', this.txtHeight);
                this.spnShpWidth = this._spinner(t, 'elx-lbl-shp-width', 'elx-spin-shp-width', this.txtWidth);
            },

            // ---- Equation -----------------------------------------------------------------
            createEquation: function () {
                var t = 'equation', tb = this.toolbar;
                this.btnEqNew = this._btn(t, 'elx-btn-eq-new', {iconCls: 'toolbar__icon btn-insertequation', caption: this.txtEquation});
                this.btnEqConvert = this._btn(t, 'elx-btn-eq-convert', {
                    iconCls: 'toolbar__icon btn-elx-eq-convert', caption: this.txtConvert,
                    menu: new Common.UI.Menu({cls: 'ppm-toolbar shifted-right', items: [
                        {caption: this.txtCurrentProfessional, iconCls: 'menu__icon btn-professional-equation', type: 'view', value: {all: false, linear: false}},
                        {caption: this.txtCurrentLinear, iconCls: 'menu__icon btn-linear-equation', type: 'view', value: {all: false, linear: true}},
                        {caption: this.txtAllProfessional, iconCls: 'menu__icon btn-professional-equation', type: 'view', value: {all: true, linear: false}},
                        {caption: this.txtAllLinear, iconCls: 'menu__icon btn-linear-equation', type: 'view', value: {all: true, linear: true}},
                        {caption: '--'},
                        {caption: 'Unicode', checkable: true, toggleGroup: 'elx-eq-input', type: 'input', value: Asc.c_oAscMathInputType.Unicode},
                        {caption: 'LaTeX', checkable: true, toggleGroup: 'elx-eq-input', type: 'input', value: Asc.c_oAscMathInputType.LaTeX}
                    ]})
                });
                this.btnEqDisplay = this._btn(t, 'elx-btn-eq-display', {iconCls: 'toolbar__icon btn-elx-eq-display', caption: this.txtChangeToDisplay});
                this.btnsEqStructures = [];
            },

            // structure buttons need the equation groups, which load with the document
            createEquationStructures: function (store) {
                var me = this, $host = this.panels.equation.find('.elx-eq-structures');
                if (this.btnsEqStructures.length || !store || !store.length) return;
                store.each(function (grp, i) {
                    if (i === 0) return; // the Symbols group goes to its own button below
                    var $slot = $('<span class="btn-slot text x-huge"></span>').appendTo($host);
                    var btn = new Common.UI.Button({
                        cls: 'btn-toolbar x-huge icon-top elx-eq-structure',
                        iconCls: 'svgicon ' + grp.get('groupIcon'),
                        caption: grp.get('groupName'),
                        lock: me._baseLock(),
                        menu: new Common.UI.Menu({
                            value: i,
                            items: [{template: _.template('<div class="menu-shape margin-left-5 elx-eq-menu" style="width:' + (grp.get('groupWidth') + 8) + 'px; ' + grp.get('groupHeightStr') + '"></div>')}]
                        }),
                        dataHint: '1', dataHintDirection: 'bottom', dataHintOffset: 'small'
                    });
                    btn.render($slot);
                    me._eqIcon(btn, grp);
                    btn.eqGroup = grp;
                    me.btnsEqStructures.push(btn);
                    me.controls.equation.push(btn);
                    me.lockedControls.push(btn);
                });
                // Symbols gallery before the structures, like Word
                var sym = store.at(0);
                if (sym) {
                    var $sg = $('<div class="group" data-group="symbols"><span class="btn-slot text x-huge"></span></div>' + SEP);
                    this.panels.equation.find('.group[data-group=structures]').before($sg);
                    var b = new Common.UI.Button({
                        cls: 'btn-toolbar x-huge icon-top elx-eq-structure',
                        iconCls: 'svgicon ' + sym.get('groupIcon'),
                        caption: sym.get('groupName'),
                        lock: me._baseLock(),
                        menu: new Common.UI.Menu({value: 0, items: [{template: _.template('<div class="menu-shape margin-left-5 elx-eq-menu" style="width:' + (sym.get('groupWidth') + 8) + 'px; ' + sym.get('groupHeightStr') + '"></div>')}]}),
                        dataHint: '1', dataHintDirection: 'bottom', dataHintOffset: 'small'
                    });
                    b.render($sg.find('.btn-slot'));
                    me._eqIcon(b, sym);
                    b.eqGroup = sym;
                    me.btnsEqStructures.unshift(b);
                    me.controls.equation.push(b);
                    me.lockedControls.push(b);
                }
                this.toolbar.clearActiveData && this.toolbar.clearActiveData('equation');
                return this.btnsEqStructures;
            },

            // the structure icons are symbols of the equation sprite (equationicons.svg)
            _eqIcon: function (btn, grp) {
                btn.$el.find('i.icon').replaceWith('<svg class="icon uni-scale"><use href="#' + grp.get('groupIcon') + '"></use></svg>');
            },

            updateMetricUnit: function () {
                _.each(this.spinners, function (spinner) {
                    spinner.setDefaultUnit(Common.Utils.Metric.getCurrentMetricName());
                    spinner.setStep(Common.Utils.Metric.getCurrentMetric() == Common.Utils.Metric.c_MetricUnits.pt ? 1 : 0.1);
                });
            },

            getButtons: function () {
                return this.lockedControls;
            },

            // ---- strings (English; translations in locale/*.json) -------------------------
            txtTabTableDesign: 'Table Design',
            txtTabTableLayout: 'Layout',
            txtTabPictureFormat: 'Picture Format',
            txtTabShapeFormat: 'Shape Format',
            txtTabEquation: 'Equation',
            txtGroupTablestyleoptions: 'Table Style Options',
            txtGroupTablestyles: 'Table Styles',
            txtGroupBorders: 'Borders',
            txtGroupTable: 'Table',
            txtGroupDraw: 'Draw',
            txtGroupRowscolumns: 'Rows & Columns',
            txtGroupMerge: 'Merge',
            txtGroupCellsize: 'Cell Size',
            txtGroupAlignment: 'Alignment',
            txtGroupData: 'Data',
            txtGroupAdjust: 'Adjust',
            txtGroupAccessibility: 'Accessibility',
            txtGroupArrange: 'Arrange',
            txtGroupSize: 'Size',
            txtGroupInsertshapes: 'Insert Shapes',
            txtGroupShapestyles: 'Shape Styles',
            txtGroupText: 'Text',
            txtGroupTools: 'Tools',
            txtGroupSymbols: 'Symbols',
            txtGroupStructures: 'Structures',
            txtHeaderRow: 'Header Row',
            txtTotalRow: 'Total Row',
            txtBandedRows: 'Banded Rows',
            txtFirstColumn: 'First Column',
            txtLastColumn: 'Last Column',
            txtBandedColumns: 'Banded Columns',
            txtShading: 'Shading',
            txtPenWeight: 'Pen Weight',
            txtPenColor: 'Pen Color',
            txtBorders: 'Borders',
            txtBorderBottom: 'Bottom Border',
            txtBorderTop: 'Top Border',
            txtBorderLeft: 'Left Border',
            txtBorderRight: 'Right Border',
            txtBorderNone: 'No Border',
            txtBorderAll: 'All Borders',
            txtBorderOutside: 'Outside Borders',
            txtBorderInside: 'Inside Borders',
            txtBorderInsideH: 'Inside Horizontal Border',
            txtBorderInsideV: 'Inside Vertical Border',
            txtBordersShading: 'Borders and Shading...',
            txtSelect: 'Select',
            txtSelectCell: 'Select Cell',
            txtSelectColumn: 'Select Column',
            txtSelectRow: 'Select Row',
            txtSelectTable: 'Select Table',
            txtViewGridlines: 'View Gridlines',
            txtProperties: 'Properties',
            txtDrawTable: 'Draw Table',
            txtEraser: 'Eraser',
            txtDelete: 'Delete',
            txtDeleteCells: 'Delete Cells...',
            txtDeleteColumns: 'Delete Columns',
            txtDeleteRows: 'Delete Rows',
            txtDeleteTable: 'Delete Table',
            txtInsertAbove: 'Insert Above',
            txtInsertBelow: 'Insert Below',
            txtInsertLeft: 'Insert Left',
            txtInsertRight: 'Insert Right',
            txtMergeCells: 'Merge Cells',
            txtSplitCells: 'Split Cells',
            txtAutoFit: 'AutoFit',
            txtAutoFitContents: 'AutoFit Contents',
            txtAutoFitWindow: 'AutoFit Window',
            txtFixedColumnWidth: 'Fixed Column Width',
            txtHeight: 'Height',
            txtWidth: 'Width',
            txtDistributeRows: 'Distribute Rows',
            txtDistributeColumns: 'Distribute Columns',
            txtAlignTopLeft: 'Align Top Left',
            txtAlignTopCenter: 'Align Top Center',
            txtAlignTopRight: 'Align Top Right',
            txtAlignCenterLeft: 'Align Center Left',
            txtAlignCenter: 'Align Center',
            txtAlignCenterRight: 'Align Center Right',
            txtAlignBottomLeft: 'Align Bottom Left',
            txtAlignBottomCenter: 'Align Bottom Center',
            txtAlignBottomRight: 'Align Bottom Right',
            txtTextDirection: 'Text Direction',
            txtDirHorizontal: 'Horizontal',
            txtDirRotateDown: 'Rotate all text 90°',
            txtDirRotateUp: 'Rotate all text 270°',
            txtCellMargins: 'Cell Margins',
            txtRepeatHeader: 'Repeat Header Rows',
            txtConvertToText: 'Convert to Text',
            txtFormula: 'Formula',
            txtPosition: 'Position',
            txtInLine: 'In Line with Text',
            txtPosTopLeft: 'Position in Top Left',
            txtPosTopCenter: 'Position in Top Center',
            txtPosTopRight: 'Position in Top Right',
            txtPosCenterLeft: 'Position in Middle Left',
            txtPosCenterCenter: 'Position in Middle Center',
            txtPosCenterRight: 'Position in Middle Right',
            txtPosBottomLeft: 'Position in Bottom Left',
            txtPosBottomCenter: 'Position in Bottom Center',
            txtPosBottomRight: 'Position in Bottom Right',
            txtMoreLayout: 'More Layout Options...',
            txtWrapText: 'Wrap Text',
            txtRotate: 'Rotate',
            txtRotateRight: 'Rotate Right 90°',
            txtRotateLeft: 'Rotate Left 90°',
            txtFlipVertical: 'Flip Vertical',
            txtFlipHorizontal: 'Flip Horizontal',
            txtMoreRotation: 'More Rotation Options...',
            txtTransparency: 'Transparency',
            txtChangePicture: 'Change Picture',
            txtFromFile: 'From a File...',
            txtFromUrl: 'From URL...',
            txtFromStorage: 'From Storage...',
            txtResetPicture: 'Reset Picture',
            txtResetPictureSize: 'Reset Picture & Size',
            txtAltText: 'Alt Text',
            txtCrop: 'Crop',
            txtFill: 'Fill',
            txtFit: 'Fit',
            txtEditShape: 'Edit Shape',
            txtChangeShape: 'Change Shape',
            txtEditPoints: 'Edit Points',
            txtShapeFill: 'Shape Fill',
            txtShapeOutline: 'Shape Outline',
            txtWeight: 'Weight',
            txtShapeEffects: 'Shape Effects',
            txtPictureBorder: 'Picture Border',
            txtPictureEffects: 'Picture Effects',
            txtGroupPicturestyles: 'Picture Styles',
            txtNoShadow: 'No Shadow',
            txtShadowOuter: 'Shadow',
            txtShadowOptions: 'Shadow Options...',
            txtAlignText: 'Align Text',
            txtAlignTop: 'Top',
            txtAlignMiddle: 'Middle',
            txtAlignBottom: 'Bottom',
            txtEquation: 'Equation',
            txtConvert: 'Convert',
            txtCurrentProfessional: 'Current - Professional',
            txtCurrentLinear: 'Current - Linear',
            txtAllProfessional: 'All - Professional',
            txtAllLinear: 'All - Linear',
            txtChangeToDisplay: 'Change to Display',
            txtChangeToInline: 'Change to Inline'
        };
    }()), DE.Views.ContextTabs || {}));
});

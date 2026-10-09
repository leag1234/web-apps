/*
 * SPDX-FileCopyrightText: 2026 ElixirOffice contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 *
 * SSE.Views.RibbonTabs: the Excel-like additions of the spreadsheet ribbon
 * (see SSE.Controllers.RibbonTabs): group labels, tab names, and the controls
 * Excel has and the stock ribbon lacks (Increase / Decrease Indent,
 * Sort & Filter, Find & Select).
 */

define([
    'core',
    'common/main/lib/component/Button',
    'common/main/lib/component/Menu',
    'common/main/lib/component/RibbonMirror'
], function () {
    'use strict';

    SSE.Views.RibbonTabs = Backbone.View.extend(_.extend({

        initialize: function (options) {
            this.toolbar = options.toolbar;
            this.mode = options.mode || {};
            this.lockedControls = [];
        },

        // groups of the panels built by other controllers: [selector, group key]
        groupTags: {
            review: [
                ['#slot-btn-spelling', 'proofing'],
                ['#slot-btn-sharing, #slot-btn-coauthmode', 'collaboration'],
                ['.slot-comment, #slot-comment-remove', 'comments'],
                ['#slot-btn-chat', 'chat'],
                ['#slot-btn-history', 'history'],
                ['#slot-btn-add-password, #slot-btn-signature', 'protect'],
                ['#slot-btn-protect-wb, #slot-btn-protect-sheet', 'protect'],
                ['#slot-chk-locked-cell', 'protect'],
                ['#slot-chk-locked-shape', 'protect'],
                ['#slot-btn-protect-range', 'protect']
            ],
            view: [
                ['#slot-btn-sheet-view', 'sheetview'],
                ['#slot-btn-view-normal', 'workbookviews'],
                ['#slot-field-zoom', 'zoom'],
                ['#slot-btn-interface-theme', 'appearance'],
                ['#slot-btn-freeze', 'window'],
                ['#slot-chk-formula', 'show'],
                ['#slot-chk-toolbar', 'interface'],
                ['#slot-btn-macros', 'macros']
            ],
            draw: [
                ['#slot-btn-draw-select, #slot-btn-draw-eraser', 'drawingtools']
            ],
            pivot: [
                ['.slot-add-pivot', 'pivottable'],
                ['#slot-btn-pivot-report-layout', 'layout'],
                ['#slot-btn-refresh-pivot', 'data'],
                ['#slot-btn-select-pivot', 'actions'],
                ['#slot-btn-calculated-items', 'calculations'],
                ['#slot-btn-expand-field', 'activefield'],
                ['#slot-chk-header-row', 'pivotstyleoptions'],
                ['#slot-field-pivot-styles', 'pivotstyles']
            ],
            tabledesign: [
                ['#slot-btn-resize-table', 'properties'],
                ['#slot-btn-remove-duplicates', 'tools'],
                ['#slot-btn-insert-slicer', 'tools'],
                ['#slot-chk-header-row', 'tablestyleoptions'],
                ['#slot-btn-alt-text', 'accessibility'],
                ['#slot-field-table-styles', 'tablestyles']
            ]
        },

        btn: function (cfg) {
            var b = new Common.UI.Button(_.extend({
                cls: 'btn-toolbar',
                dataHint: '1',
                dataHintDirection: 'bottom',
                dataHintOffset: 'small',
                lock: [Common.enumLock.lostConnect, Common.enumLock.coAuth, Common.enumLock.disableOnStart]
            }, cfg));
            this.lockedControls.push(b);
            return b;
        },

        // created on Toolbar 'render:before', rendered on 'render:after'
        createControls: function () {
            var me = this, _set = Common.enumLock,
                cellLock = [_set.editCell, _set.selChart, _set.selChartText, _set.selShape, _set.selShapeText, _set.selImage, _set.selSlicer,
                            _set.lostConnect, _set.coAuth, _set['FormatCells'], _set.editVisibleArea, _set.disableOnStart];
            this.btnIndentDec = this.btn({
                id: 'id-toolbar-btn-elx-indent-dec', iconCls: 'toolbar__icon btn-decoffset', hint: this.tipDecIndent, lock: cellLock
            });
            this.btnIndentInc = this.btn({
                id: 'id-toolbar-btn-elx-indent-inc', iconCls: 'toolbar__icon btn-incoffset', hint: this.tipIncIndent, lock: cellLock
            });
            this.btnSortFilter = this.btn({
                id: 'id-toolbar-btn-elx-sort-filter', cls: 'btn-toolbar x-huge icon-top', iconCls: 'toolbar__icon btn-big-elx-sort-filter',
                caption: this.txtSortFilter, hint: this.tipSortFilter, lock: [_set.lostConnect, _set.coAuth, _set.disableOnStart],
                menu: new Common.UI.Menu({
                    items: [
                        {caption: this.txtSortAZ, value: 'sort-down', iconCls: 'menu__icon btn-sort-down'},
                        {caption: this.txtSortZA, value: 'sort-up', iconCls: 'menu__icon btn-sort-up'},
                        {caption: this.txtCustomSort, value: 'sort-custom', iconCls: 'menu__icon btn-custom-sort'},
                        {caption: '--'},
                        {caption: this.txtFilter, value: 'set-filter', iconCls: 'menu__icon btn-autofilter', checkable: true},
                        {caption: this.txtClearFilter, value: 'clear-filter', iconCls: 'menu__icon btn-clear-filter'}
                    ]
                })
            });
            // Insert > Charts: one gallery per family of chart types, as in Excel
            var chartLock = [_set.editCell, _set.lostConnect, _set.coAuth, _set.coAuthText, _set['Objects'], _set.disableOnStart];
            var gallery = function (key, hint, icon, groups) {
                var b = me.btn({
                    id: 'id-toolbar-btn-elx-chart-' + key, iconCls: 'toolbar__icon btn-insertchart', hint: hint, lock: chartLock,
                    menu: new Common.UI.Menu({
                        style: 'width: 364px; padding-top: 12px;',
                        items: [
                            {template: _.template('<div id="id-elx-chart-gallery-' + key + '" class="menu-insertchart"></div>')},
                            {caption: '--'},
                            {caption: me.txtAllCharts, value: 'all-charts'}
                        ]
                    })
                });
                b.elxKey = key; b.elxIcon = icon; b.elxGroups = groups;
                return b;
            };
            this.btnsCharts = [
                gallery('column', this.tipChartColumn, 'column-normal', ['menu-chart-group-bar', 'menu-chart-group-hbar']),
                gallery('line', this.tipChartLine, 'line-normal', ['menu-chart-group-line', 'menu-chart-group-area']),
                gallery('pie', this.tipChartPie, 'pie-normal', ['menu-chart-group-pie']),
                gallery('scatter', this.tipChartScatter, 'point-normal', ['menu-chart-group-scatter']),
                gallery('stock', this.tipChartStock, 'stock-normal', ['menu-chart-group-stock', 'menu-chart-group-radar']),
                gallery('combo', this.tipChartCombo, 'column-normal', ['menu-chart-group-combo'])
            ];
            this.btnFindSelect = this.btn({
                id: 'id-toolbar-btn-elx-find-select', cls: 'btn-toolbar x-huge icon-top', iconCls: 'toolbar__icon btn-big-elx-find-select',
                caption: this.txtFindSelect, hint: this.tipFindSelect, lock: [_set.disableOnStart],
                menu: new Common.UI.Menu({
                    items: [
                        {caption: this.txtFind, value: 'find', iconCls: 'menu__icon btn-menu-search'},
                        {caption: this.txtReplace, value: 'replace', iconCls: 'menu__icon btn-replace'},
                        {caption: this.txtGoTo, value: 'goto'},
                        {caption: '--'},
                        {caption: this.txtSelectAll, value: 'selectall', iconCls: 'menu__icon btn-select-all'}
                    ]
                })
            });
        },

        renderControls: function ($host) {
            var put = function (sel, b) {
                var $slot = $host.find(sel);
                $slot.length && b.render($slot);
            };
            put('#slot-elx-indent-dec', this.btnIndentDec);
            put('#slot-elx-indent-inc', this.btnIndentInc);
            put('#slot-elx-sort-filter', this.btnSortFilter);
            put('#slot-elx-find-select', this.btnFindSelect);
            _.each(this.btnsCharts, function (b) {
                put('#slot-elx-chart-' + b.elxKey, b);
                // the colourful preview of the first type of the family as the icon
                b.cmpEl && b.cmpEl.find('button').first().find('.icon, .btn-icon').first().replaceWith(
                    '<svg class="elx-chart-icon" width="18" height="18" aria-hidden="true"><use xlink:href="#chart-' + b.elxIcon + '"></use></svg>');
            });
        },

        // ---- strings ----------------------------------------------------------------
        txtGroupClipboard: 'Clipboard',
        txtGroupFont: 'Font',
        txtGroupAlignment: 'Alignment',
        txtGroupNumber: 'Number',
        txtGroupStyles: 'Styles',
        txtGroupCellstyles: 'Cell Styles',
        txtGroupCells: 'Cells',
        txtGroupEditing: 'Editing',
        txtGroupTables: 'Tables',
        txtGroupIllustrations: 'Illustrations',
        txtGroupCharts: 'Charts',
        txtGroupSparklines: 'Sparklines',
        txtGroupFilters: 'Filters',
        txtGroupLinks: 'Links',
        txtGroupComments: 'Comments',
        txtGroupText: 'Text',
        txtGroupSymbols: 'Symbols',
        txtGroupCellcontrols: 'Cell Controls',
        txtGroupThemes: 'Themes',
        txtGroupPagesetup: 'Page Setup',
        txtGroupScaletofit: 'Scale to Fit',
        txtGroupSheetoptions: 'Sheet Options',
        txtGroupArrange: 'Arrange',
        txtGroupFunctionlibrary: 'Function Library',
        txtGroupDefinednames: 'Defined Names',
        txtGroupFormulaauditing: 'Formula Auditing',
        txtGroupCalculation: 'Calculation',
        txtGroupGettransform: 'Get & Transform Data',
        txtGroupQueries: 'Queries & Connections',
        txtGroupSortfilter: 'Sort & Filter',
        txtGroupDatatools: 'Data Tools',
        txtGroupForecast: 'Forecast',
        txtGroupOutline: 'Outline',
        txtGroupAnalyze: 'Analyze',
        txtGroupProofing: 'Proofing',
        txtGroupCollaboration: 'Collaboration',
        txtGroupChat: 'Chat',
        txtGroupHistory: 'History',
        txtGroupProtect: 'Protect',
        txtGroupSheetview: 'Sheet View',
        txtGroupWorkbookviews: 'Workbook Views',
        txtGroupZoom: 'Zoom',
        txtGroupAppearance: 'Appearance',
        txtGroupWindow: 'Window',
        txtGroupShow: 'Show',
        txtGroupInterface: 'Interface',
        txtGroupMacros: 'Macros',
        txtGroupDrawingtools: 'Drawing Tools',
        txtGroupPivottable: 'PivotTable',
        txtGroupLayout: 'Layout',
        txtGroupData: 'Data',
        txtGroupActions: 'Actions',
        txtGroupCalculations: 'Calculations',
        txtGroupActivefield: 'Active Field',
        txtGroupPivotstyleoptions: 'PivotTable Style Options',
        txtGroupPivotstyles: 'PivotTable Styles',
        txtGroupProperties: 'Properties',
        txtGroupTools: 'Tools',
        txtGroupTablestyleoptions: 'Table Style Options',
        txtGroupAccessibility: 'Accessibility',
        txtGroupTablestyles: 'Table Styles',
        txtGroupType: 'Type',
        txtGroupCurrentselection: 'Current Selection',
        txtGroupInsertshapes: 'Insert Shapes',
        txtGroupShapestyles: 'Shape Styles',
        txtGroupPicturestyles: 'Picture Styles',
        tipGroupLauncher: 'More {0} options',

        txtTabPageLayout: 'Page Layout',
        txtTabFormulas: 'Formulas',
        txtTabReview: 'Review',
        txtTabPivot: 'PivotTable',

        txtPaste: 'Paste',
        txtCondFormat: 'Conditional Formatting',
        txtFormatAsTable: 'Format as Table',
        txtInsert: 'Insert',
        txtDelete: 'Delete',
        txtFormat: 'Format',
        txtWrapText: 'Wrap Text',
        txtMergeCenter: 'Merge & Center',
        txtAutoSum: 'AutoSum',
        txtFill: 'Fill',
        txtClear: 'Clear',
        txtSortFilter: 'Sort & Filter',
        tipSortFilter: 'Sort & Filter',
        txtSortAZ: 'Sort A to Z',
        txtSortZA: 'Sort Z to A',
        txtCustomSort: 'Custom Sort...',
        txtFilter: 'Filter',
        txtClearFilter: 'Clear',
        txtFindSelect: 'Find & Select',
        tipFindSelect: 'Find & Select',
        txtFind: 'Find...',
        txtReplace: 'Replace...',
        txtGoTo: 'Go To...',
        txtSelectAll: 'Select All',
        tipDecIndent: 'Decrease Indent',
        tipIncIndent: 'Increase Indent',
        txtAllCharts: 'All Charts...',
        tipChartColumn: 'Insert Column or Bar Chart',
        tipChartLine: 'Insert Line or Area Chart',
        tipChartPie: 'Insert Pie or Doughnut Chart',
        tipChartScatter: 'Insert Scatter (X, Y) Chart',
        tipChartStock: 'Insert Stock or Radar Chart',
        tipChartCombo: 'Insert Combo Chart',
        txtCollapseRibbon: 'Collapse the Ribbon',
        txtPinRibbon: 'Pin the ribbon'

    }, SSE.Views.RibbonTabs || {}));
});

/*
 * SPDX-FileCopyrightText: 2026 ElixirOffice contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 *
 * PE.Controllers.SlideContextTabs: PowerPoint-like contextual ribbon tabs of
 * the presentation editor: Shape Format, Picture Format, Table Design, Table
 * Layout, Chart Design. They appear at the right of the tab row, in the
 * accent colour, while their object is selected (asc_onFocusObject), and
 * disappear with it; inserting an object opens its tab, like PowerPoint.
 *
 * The commands use the same API calls as the right panel (ShapeApply,
 * ImgApply, tblApply...) or are twins of the stock ribbon buttons
 * (Common.UI.RibbonMirror); the right panel stays as the detailed "Format" pane
 * (group launchers).
 */
define([
    'core',
    'common/main/lib/component/Button',
    'common/main/lib/component/RibbonMirror',
    'common/main/lib/component/CheckBox',
    'common/main/lib/component/MetricSpinner',
    'common/main/lib/component/ColorButton',
    'common/main/lib/component/ComboBorderSize',
    'common/main/lib/component/WordArtRibbon',
    'common/main/lib/component/ChartRibbon'
], function () {
    'use strict';

    var TABS = ['shapeformat', 'pictureformat', 'tabledesign', 'tablelayout', 'chartdesign'];
    var KEYTIPS = { shapeformat: 'JD', pictureformat: 'JP', tabledesign: 'JT', tablelayout: 'JL', chartdesign: 'JC' };

    var CELL_MARGINS = { none: [0, 0, 0, 0], narrow: [1.27, 1.27, 1.27, 1.27], moderate: [1.905, 1.27, 1.905, 1.27], wide: [5.08, 1.27, 5.08, 1.27] };

    var STRINGS = {
        txtTabShapeFormat: 'Shape Format', txtTabPictureFormat: 'Picture Format', txtTabTableDesign: 'Table Design',
        txtTabTableLayout: 'Layout', txtTabChartDesign: 'Chart Design',
        txtGroupShapeStyles: 'Shape Styles', txtGroupArrange: 'Arrange', txtGroupSize: 'Size', txtGroupAdjust: 'Adjust',
        txtGroupPictureStyles: 'Picture Styles', txtGroupStyleOptions: 'Table Style Options', txtGroupTableStyles: 'Table Styles',
        txtGroupRowsCols: 'Rows & Columns', txtGroupMerge: 'Merge', txtGroupTable: 'Table', txtGroupAlignment: 'Alignment',
        txtGroupType: 'Type', txtGroupData: 'Data', txtGroupInsertShapes: 'Insert Shapes',
        txtShapeFill: 'Shape Fill', txtShapeOutline: 'Shape Outline', txtFormatPane: 'Format Pane',
        txtNoFill: 'No Fill', txtNoOutline: 'No Outline', txtPictureBorder: 'Picture Border',
        txtChangePicture: 'Change Picture', txtFromFile: 'From a File...', txtFromUrl: 'From URL...', txtFromStorage: 'From Storage',
        txtResetPicture: 'Reset Picture', txtCrop: 'Crop', txtFill: 'Fill', txtFit: 'Fit',
        txtBringForward: 'Bring Forward', txtSendBackward: 'Send Backward', txtGroup: 'Group', txtRotate: 'Rotate',
        txtRotateRight: 'Rotate Right 90°', txtRotateLeft: 'Rotate Left 90°', txtFlipV: 'Flip Vertical', txtFlipH: 'Flip Horizontal',
        txtHeight: 'Height', txtWidth: 'Width',
        txtHeaderRow: 'Header Row', txtTotalRow: 'Total Row', txtBandedRows: 'Banded Rows', txtFirstColumn: 'First Column',
        txtLastColumn: 'Last Column', txtBandedColumns: 'Banded Columns', txtShading: 'Shading',
        txtSelect: 'Select', txtSelectCell: 'Select Cell', txtSelectColumn: 'Select Column', txtSelectRow: 'Select Row', txtSelectTable: 'Select Table',
        txtDelete: 'Delete', txtDeleteColumns: 'Delete Columns', txtDeleteRows: 'Delete Rows', txtDeleteTable: 'Delete Table',
        txtInsertAbove: 'Insert Above', txtInsertBelow: 'Insert Below', txtInsertLeft: 'Insert Left', txtInsertRight: 'Insert Right',
        txtMergeCells: 'Merge Cells', txtSplitCells: 'Split Cells', txtDistRows: 'Distribute Rows', txtDistCols: 'Distribute Columns',
        txtAlignTop: 'Align Top', txtAlignMiddle: 'Center Vertically', txtAlignBottom: 'Align Bottom',
        txtChangeChartType: 'Change Chart Type', txtEditData: 'Edit Data', txtProperties: 'Properties',
        txtGroupDrawBorders: 'Draw Borders', txtGroupCellSize: 'Cell Size', txtGroupWordArt: 'WordArt Styles',
        txtPenWeight: 'Pen Weight', txtPenColor: 'Pen Color', txtBorders: 'Borders',
        txtBorderBottom: 'Bottom Border', txtBorderTop: 'Top Border', txtBorderLeft: 'Left Border', txtBorderRight: 'Right Border',
        txtBorderNone: 'No Border', txtBorderAll: 'All Borders', txtBorderOutside: 'Outside Borders', txtBorderInside: 'Inside Borders',
        txtBorderInsideH: 'Inside Horizontal Border', txtBorderInsideV: 'Inside Vertical Border', txtTableProperties: 'Table Properties...',
        txtTextDirection: 'Text Direction', txtDirHorizontal: 'Horizontal', txtDirRotateDown: 'Rotate all text 90°', txtDirRotateUp: 'Rotate all text 270°',
        txtCellMargins: 'Cell Margins', txtMarginsNormal: 'Normal', txtMarginsNone: 'None', txtMarginsNarrow: 'Narrow',
        txtMarginsModerate: 'Moderate', txtMarginsWide: 'Wide', txtMarginsCustom: 'Custom Margins...',
        txtWordArtQuickStyles: 'Quick Styles', txtTextFill: 'Text Fill', txtTextOutline: 'Text Outline', txtTextEffects: 'Text Effects',
        txtTransform: 'Transform', txtWeight: 'Weight',
        txtGroupChartlayouts: 'Chart Layouts', txtGroupChartstyles: 'Chart Styles', txtAddChartElement: 'Add Chart Element',
        txtQuickLayout: 'Quick Layout', txtChangeColors: 'Change Colors', txtChartStyles: 'Chart Styles', txtLayout: 'Layout',
        txtColorful: 'Colorful', txtMonochromatic: 'Monochromatic'
    };

    PE.Views.SlideContextTabs = Backbone.View.extend(_.extend({ initialize: function () {} }, STRINGS, PE.Views.SlideContextTabs || {}));
    PE.Views.SlideContextTabs.STRINGS = STRINGS;

    PE.Controllers.SlideContextTabs = Backbone.Controller.extend({

        initialize: function () {
            var me = this;
            this._state = { visible: {}, lastTab: 'home' };
            this.addListeners({
                'Toolbar': {
                    'render:after': function (toolbar) { me.onToolbarRender(toolbar); },
                    'tab:active': function (tab) { if (TABS.indexOf(tab) < 0) me._state.lastTab = tab; }
                }
            });
            Common.NotificationCenter.on('app:ready', this.onAppReady.bind(this));
        },

        L: function (k) { return this.strings[k]; },

        onToolbarRender: function (tb) {
            var main = this.getApplication().getController('Main'), mode = (main && main.appOptions) || {};
            if (!mode.isEdit || this.tb) return;
            this.tb = tb;
            this.strings = new PE.Views.SlideContextTabs();
            this.controls = [];
            var me = this, s = this.strings;
            var caps = { shapeformat: s.txtTabShapeFormat, pictureformat: s.txtTabPictureFormat, tabledesign: s.txtTabTableDesign,
                         tablelayout: s.txtTabTableLayout, chartdesign: s.txtTabChartDesign };
            _.each(TABS, function (t) {
                var $panel = $('<section class="panel" data-tab="' + t + '" role="tabpanel" aria-labelledby="' + t + '"></section>');
                tb.addTab({ action: t, caption: caps[t], extcls: 'canedit contextual', layoutname: 'toolbar-' + t, dataHintTitle: KEYTIPS[t] },  // KeyTips: PE.Controllers.PowerPointKeyTips
                    $panel, tb.getLastTabIdx());
                me['build_' + t]($panel);
                tb.setVisible(t, false);
                me._state.visible[t] = false;
            });
            Array.prototype.push.apply(tb.lockControls, this.controls);
            Common.Utils.lockControls(Common.enumLock.disableOnStart, true, { array: this.controls });
            // labels of the new groups (Common.UI.RibbonGroupLabels reads tb.txtGroup<Key>)
            _.each(STRINGS, function (v, k) { if (/^txtGroup/.test(k)) tb[k] = s[k]; });
            setTimeout(function () { me.arrangeTabs(); }, 0);
        },

        // contextual tabs at the right end of the tab row
        arrangeTabs: function () {
            var tb = this.tb, $ul = tb.$tabs.first().parent();
            _.each(TABS, function (t) { $ul.find('> li > a[data-tab=' + t + ']').parent().addClass('contextual').appendTo($ul); });
            tb.$tabs = $ul.find('> li');
        },

        onAppReady: function () {
            if (!this.tb) return;
            var me = this;
            this.api = this.getApplication().getController('Viewport').getApi();
            Common.Utils.lockControls(Common.enumLock.disableOnStart, false, { array: this.controls });
            this.api.asc_registerCallback('asc_onFocusObject', _.bind(this.onFocusObject, this));
            this.api.asc_registerCallback('asc_onSendThemeColors', function () { setTimeout(function () { me.updateColors(); }, 0); });
            this.api.asc_registerCallback('asc_onCoAuthoringDisconnect', function () { me.hideAll(); });
            this.wrapInsertCommands(this.api);
            this.updateColors();
            Common.NotificationCenter.on('settings:unitschanged', function () { me.updateMetric(); });
            this.updateMetric();
            this.arrangeTabs();
        },

        // ---- builders -------------------------------------------------------------------
        // groups: [[key, html], ...] -> panel markup with separators
        layout: function ($panel, groups) {
            $panel.html(groups.map(function (g) {
                return '<div class="group' + (g[2] ? ' ' + g[2] : '') + '" data-group="' + g[0] + '"' + (g[3] ? ' data-launcher="' + g[3] + '"' : '') + '>' + g[1] + '</div>';
            }).join('<div class="separator long"></div>'));
        },
        big: function (id) { return '<span class="btn-slot text x-huge" id="' + id + '"></span>'; },
        smallRows: function (ids) {
            return ids.map(function (row) {
                return '<div class="elset">' + [].concat(row).map(function (id) { return '<span class="btn-slot text" id="' + id + '"></span>'; }).join('') + '</div>';
            }).join('');
        },

        button: function ($panel, id, opts, small) {
            var _set = Common.enumLock;
            var b = new Common.UI.Button(_.extend({
                cls: small ? 'btn-toolbar' : 'btn-toolbar x-huge icon-top',
                lock: [_set.lostConnect, _set.noSlides, _set.disableOnStart],
                dataHint: '1', dataHintDirection: small ? 'left' : 'bottom', dataHintOffset: 'small'
            }, opts));
            b.render($panel.find('#' + id));
            this.controls.push(b);
            return b;
        },

        colorButton: function ($panel, id, opts, onColor) {
            var _set = Common.enumLock, me = this;
            var b = new Common.UI.ButtonColored(_.extend({
                cls: 'btn-toolbar', split: true, menu: true,
                lock: [_set.lostConnect, _set.noSlides, _set.disableOnStart],
                dataHint: '1', dataHintDirection: 'left', dataHintOffset: 'small'
            }, opts));
            b.render($panel.find('#' + id));
            b.setMenu();
            b.picker = b.getPicker();
            b.on('color:select', function (btn, color) { onColor.call(me, color); });
            b.on('click', function (btn) { onColor.call(me, btn.currentColor || btn.color); });
            this.controls.push(b);
            (this.colorButtons || (this.colorButtons = [])).push(b);
            return b;
        },

        spinner: function ($panel, id, onChange) {
            var me = this;
            var s = new Common.UI.MetricSpinner({
                el: $panel.find('#' + id), step: .1, width: 78, defaultUnit: 'cm', value: '0 cm', maxValue: 508, minValue: 0,
                dataHint: '1', dataHintDirection: 'left', dataHintOffset: 'small'
            });
            s.on('change', function (field) { me._noApply || onChange.call(me, Common.Utils.Metric.fnRecalcToMM(field.getNumberValue())); });
            s.on('inputleave', function () { Common.NotificationCenter.trigger('edit:complete', me.tb); });
            (this.spinners || (this.spinners = [])).push(s);
            this.controls.push(s);
            return s;
        },

        sizeHtml: function (pfx) {
            var s = this.strings;
            return '<div class="elset font-normal"><label class="elx-spin-label">' + s.txtHeight + '</label><div class="btn-slot" id="' + pfx + '-height"></div></div>' +
                   '<div class="elset font-normal"><label class="elx-spin-label">' + s.txtWidth + '</label><div class="btn-slot" id="' + pfx + '-width"></div></div>';
        },

        // Arrange: twins of the Home > Drawing commands, as in PowerPoint's Format tabs
        arrange: function ($panel, pfx) {
            var tb = this.tb, me = this;
            var twin = function (id, src, caption) {
                var t = new Common.UI.RibbonMirror(function () { return tb[src]; }, { caption: caption, cls: 'btn-toolbar' });
                t.render($panel.find('#' + id));
                t.mirrorAttach();
                me.controls.push(t);
                return t;
            };
            twin(pfx + '-arrange', 'btnShapeArrange', this.strings.txtGroupArrange);
            twin(pfx + '-align', 'btnShapeAlign', this.L('txtAlignObjects') || PE.Views.SlideRibbon.prototype.txtAlign);
        },
        arrangeHtml: function (pfx) { return this.smallRows([pfx + '-arrange', pfx + '-align']); },

        // WordArt Styles group (Common.UI.WordArtRibbon, shared with the document editor)
        wordArt: function ($p, pfx) {
            var me = this, s = this.strings, _set = Common.enumLock;
            this.wa = Common.UI.WordArtRibbon.create({
                strings: s,
                ids: { styles: pfx + '-wa-styles', fill: pfx + '-wa-fill', outline: pfx + '-wa-outline', effects: pfx + '-wa-effects' },
                lock: [_set.lostConnect, _set.noSlides, _set.disableOnStart],
                add: function (id, cmp) { cmp.render($p.find('#' + id)); me.controls.push(cmp); },
                apply: function (fn) { if (!me.api) return; var props = new Asc.asc_CShapeProperty(); fn(props); me.api.ShapeApply(props); },
                api: function () { return me.api; },
                collection: function () { return PE.getCollection('Common.Collections.TextArt'); },
                fillTextArt: function () { me.api && PE.getController('Main').fillTextArt(me.api.asc_getTextArtPreviews()); },
                done: function () { me.done(); }
            });
        },

        // Shape Format: Shape Styles | WordArt Styles | Arrange | Size
        build_shapeformat: function ($p) {
            var s = this.strings, me = this;
            this.layout($p, [
                ['shapeStyles', this.smallRows(['sf-fill', 'sf-outline']) , 'small', 'shape'],
                ['', this.big('sf-pane')],
                ['wordArt', this.big('sf-wa-styles') + '<div class="elset-column">' + this.smallRows(['sf-wa-fill', 'sf-wa-outline', 'sf-wa-effects']) + '</div>', 'small elx-three-rows'],
                ['arrange', this.arrangeHtml('sf'), 'small'],
                ['size', this.sizeHtml('sf'), 'small elx-size', 'shapesize']
            ]);
            $p.find('.separator').eq(0).addClass('invisible');
            this.sfFill = this.colorButton($p, 'sf-fill', { iconCls: 'toolbar__icon btn-elx-shape-fill', caption: s.txtShapeFill }, function (c) { me.applyShapeFill(c); });
            this.sfOutline = this.colorButton($p, 'sf-outline', { iconCls: 'toolbar__icon btn-elx-shape-outline', caption: s.txtShapeOutline }, function (c) { me.applyShapeStroke(c); });
            this.button($p, 'sf-pane', { iconCls: 'toolbar__icon btn-elx-pane', caption: s.txtFormatPane })
                .on('click', function () { me.rightPanel(Asc.c_oAscTypeSelectElement.Shape); });
            this.wordArt($p, 'sf');
            this.arrange($p, 'sf');
            this.sfH = this.spinner($p, 'sf-height', function (mm) { me.applySize('shape', null, mm); });
            this.sfW = this.spinner($p, 'sf-width', function (mm) { me.applySize('shape', mm, null); });
        },

        // Picture Format: Adjust | Picture Styles | Arrange | Size (Crop)
        build_pictureformat: function ($p) {
            var s = this.strings, me = this;
            this.layout($p, [
                ['adjust', this.big('pf-change') + this.big('pf-reset')],
                ['pictureStyles', this.smallRows(['pf-border']), 'small', 'picture'],
                ['arrange', this.arrangeHtml('pf'), 'small'],
                ['size', this.big('pf-crop') + '<div class="elset-column">' + this.sizeHtml('pf') + '</div>', '', 'picturesize']
            ]);
            var change = this.button($p, 'pf-change', { iconCls: 'toolbar__icon btn-elx-change-picture', caption: s.txtChangePicture, menu: new Common.UI.Menu({
                items: [{ caption: s.txtFromFile, value: 0 }, { caption: s.txtFromUrl, value: 1 }, { caption: s.txtFromStorage, value: 2 }] }) });
            change.menu.on('item:click', function (menu, item) { me.changePicture(item.value); });
            this.button($p, 'pf-reset', { iconCls: 'toolbar__icon btn-elx-reset-picture', caption: s.txtResetPicture })
                .on('click', function () { me.resetPicture(); });
            this.pfBorder = this.colorButton($p, 'pf-border', { iconCls: 'toolbar__icon btn-elx-shape-outline', caption: s.txtPictureBorder }, function (c) { me.applyShapeStroke(c); });
            this.arrange($p, 'pf');
            this.pfCrop = this.button($p, 'pf-crop', { iconCls: 'toolbar__icon btn-elx-crop', caption: s.txtCrop, split: true, enableToggle: true,
                menu: new Common.UI.Menu({ items: [{ caption: s.txtCrop, value: 0, checkable: true }, { caption: s.txtFill, value: 1 }, { caption: s.txtFit, value: 2 }] }) });
            this.pfCrop.on('click', function (btn) { me.crop(btn.pressed ? 0 : -1); });
            this.pfCrop.menu.on('item:click', function (menu, item) { me.crop(item.value); });
            this.pfH = this.spinner($p, 'pf-height', function (mm) { me.applySize('image', null, mm); });
            this.pfW = this.spinner($p, 'pf-width', function (mm) { me.applySize('image', mm, null); });
        },

        // Table Design: Table Style Options | Table Styles (Shading, Format pane)
        build_tabledesign: function ($p) {
            var s = this.strings, me = this;
            var chk = function (id) { return '<span class="btn-slot text" id="' + id + '"></span>'; };
            this.layout($p, [
                ['styleOptions', '<div class="elset-column"><div class="elset">' + chk('td-header') + '</div><div class="elset">' + chk('td-total') + '</div><div class="elset">' + chk('td-bandrows') + '</div></div>' +
                    '<div class="elset-column"><div class="elset">' + chk('td-first') + '</div><div class="elset">' + chk('td-last') + '</div><div class="elset">' + chk('td-bandcols') + '</div></div>', 'small elx-three-rows'],
                ['tableStyles', this.smallRows(['td-shading']) , 'small', 'table'],
                ['drawBorders', '<div class="elset-column"><div class="elset"><span class="btn-slot" id="td-bsize"></span></div><div class="elset"><span class="btn-slot text" id="td-pen"></span></div></div>' + this.big('td-borders'), 'small elx-borders'],
                ['', this.big('td-pane')]
            ]);
            $p.find('.separator').eq(2).addClass('invisible');
            this.tdChecks = {};
            var looks = [['td-header', 'txtHeaderRow', 0], ['td-total', 'txtTotalRow', 1], ['td-bandrows', 'txtBandedRows', 2],
                         ['td-first', 'txtFirstColumn', 3], ['td-last', 'txtLastColumn', 4], ['td-bandcols', 'txtBandedColumns', 5]];
            _.each(looks, function (l) {
                var c = new Common.UI.CheckBox({ el: $p.find('#' + l[0]), labelText: s[l[1]], value: false, dataHint: '1', dataHintDirection: 'left', dataHintOffset: 'small' });
                c.on('change', function (field) { me._noApply || me.applyTableLook(l[2], field.getValue() === 'checked'); });
                me.tdChecks[l[2]] = c;
                me.controls.push(c);
            });
            this.drawBorders($p);
            this.tdShading = this.colorButton($p, 'td-shading', { iconCls: 'toolbar__icon btn-elx-shape-fill', caption: s.txtShading }, function (c) { me.applyCellShading(c); });
            this.button($p, 'td-pane', { iconCls: 'toolbar__icon btn-elx-pane', caption: s.txtFormatPane })
                .on('click', function () { me.rightPanel(Asc.c_oAscTypeSelectElement.Table); });
        },

        // Draw Borders: pen weight and colour, Borders presets (the Word controls on the presentation table API)
        drawBorders: function ($p) {
            var s = this.strings, me = this, _set = Common.enumLock, lock = [_set.lostConnect, _set.noSlides, _set.disableOnStart];
            var hint = { dataHint: '1', dataHintDirection: 'bottom', dataHintOffset: 'big' };
            this.tdBorderSize = new Common.UI.ComboBorderSize(_.extend({ style: 'width: 93px;', lock: lock, ariaLabel: s.txtPenWeight }, hint));
            this.tdBorderSize.render($p.find('#td-bsize'));
            this._borderSize = this.tdBorderSize.store.at(1).get('value');
            this.tdBorderSize.setValue(this._borderSize);
            this.tdBorderSize.on('selected', function (combo, record) { me._borderSize = record.value; });
            this.controls.push(this.tdBorderSize);

            this.tdPen = new Common.UI.ButtonColored({
                cls: 'btn-toolbar', iconCls: 'toolbar__icon btn-elx-pen-color', caption: s.txtPenColor, menu: true, auto: true, color: 'auto',
                eyeDropper: true, lock: lock, dataHint: '1', dataHintDirection: 'left', dataHintOffset: 'small'
            });
            this.tdPen.render($p.find('#td-pen'));
            this.tdPen.setMenu();
            this.tdPen.getPicker();
            this.controls.push(this.tdPen);

            var bi = function (value, cls, caption) { return { caption: caption, iconCls: 'menu__icon ' + cls, value: value }; };
            this.tdBorders = new Common.UI.Button({
                cls: 'btn-toolbar x-huge icon-top', iconCls: 'toolbar__icon btn-elx-borders', caption: s.txtBorders, split: true, lock: lock,
                dataHint: '1', dataHintDirection: 'bottom', dataHintOffset: 'small',
                menu: new Common.UI.Menu({ cls: 'shifted-right', items: [
                    bi('b', 'btn-border-bottom', s.txtBorderBottom), bi('t', 'btn-border-top', s.txtBorderTop),
                    bi('l', 'btn-border-left', s.txtBorderLeft), bi('r', 'btn-border-right', s.txtBorderRight),
                    { caption: '--' },
                    bi('', 'btn-border-no', s.txtBorderNone), bi('lrtbcm', 'btn-border-all', s.txtBorderAll),
                    bi('lrtb', 'btn-border-out', s.txtBorderOutside), bi('cm', 'btn-border-inside', s.txtBorderInside),
                    bi('m', 'btn-border-insidehor', s.txtBorderInsideH), bi('c', 'btn-border-insidevert', s.txtBorderInsideV),
                    { caption: '--' },
                    { caption: s.txtTableProperties, value: 'advanced' }
                ] })
            });
            this.tdBorders.render($p.find('#td-borders'));
            this.tdBorders.options.lastBorder = 'b';
            this.controls.push(this.tdBorders);
            this.tdBorders.menu.on('item:click', function (menu, item) {
                if (item.value === 'advanced') { me.tableProperties(); return; }
                me.tdBorders.options.lastBorder = item.value;
                me.applyCellBorders(item.value);
            });
            this.tdBorders.on('click', function () { me.applyCellBorders(me.tdBorders.options.lastBorder); });
        },

        // Layout: Table (Select) | Rows & Columns | Merge | Cell Size | Alignment
        build_tablelayout: function ($p) {
            var s = this.strings, me = this, api = function () { return me.api; };
            this.layout($p, [
                ['table', this.big('tl-select')],
                ['rowsCols', this.big('tl-delete') + this.big('tl-above') + this.big('tl-below') + this.big('tl-left') + this.big('tl-right')],
                ['merge', this.big('tl-merge') + this.big('tl-split')],
                ['cellSize', '<div class="elset-column">' + this.sizeHtml('tl') + '</div><div class="elset-column">' + this.smallRows(['tl-distrows', 'tl-distcols']) + '</div>', 'small elx-size'],
                ['alignment', '<div class="elset-column">' + this.smallRows(['tl-top', 'tl-middle', 'tl-bottom']) + '</div><div class="elset-column">' + this.smallRows(['tl-direction', 'tl-margins']) + '</div>', 'small elx-three-rows']
            ]);
            var sel = this.button($p, 'tl-select', { iconCls: 'toolbar__icon btn-elx-tbl-select', caption: s.txtSelect, menu: new Common.UI.Menu({
                items: [{ caption: s.txtSelectCell, value: 'cell' }, { caption: s.txtSelectColumn, value: 'col' }, { caption: s.txtSelectRow, value: 'row' }, { caption: s.txtSelectTable, value: 'table' }] }) });
            sel.menu.on('item:click', function (m, it) { var a = api(); a && ({ cell: a.selectCell, col: a.selectColumn, row: a.selectRow, table: a.selectTable })[it.value].call(a); me.done(); });
            var del = this.button($p, 'tl-delete', { iconCls: 'toolbar__icon btn-elx-tbl-delete', caption: s.txtDelete, menu: new Common.UI.Menu({
                items: [{ caption: s.txtDeleteColumns, value: 'col' }, { caption: s.txtDeleteRows, value: 'row' }, { caption: s.txtDeleteTable, value: 'table' }] }) });
            del.menu.on('item:click', function (m, it) { var a = api(); a && ({ col: a.remColumn, row: a.remRow, table: a.remTable })[it.value].call(a); me.done(); });
            var cmd = function (id, icon, cap, fn) { me.button($p, id, { iconCls: 'toolbar__icon ' + icon, caption: cap }).on('click', function () { var a = api(); a && fn(a); me.done(); }); };
            cmd('tl-above', 'btn-elx-tbl-above', s.txtInsertAbove, function (a) { a.addRowAbove(); });
            cmd('tl-below', 'btn-elx-tbl-below', s.txtInsertBelow, function (a) { a.addRowBelow(); });
            cmd('tl-left', 'btn-elx-tbl-left', s.txtInsertLeft, function (a) { a.addColumnLeft(); });
            cmd('tl-right', 'btn-elx-tbl-right', s.txtInsertRight, function (a) { a.addColumnRight(); });
            cmd('tl-merge', 'btn-elx-tbl-merge', s.txtMergeCells, function (a) { a.MergeCells(); });
            this.button($p, 'tl-split', { iconCls: 'toolbar__icon btn-elx-tbl-split', caption: s.txtSplitCells }).on('click', function () { me.splitCells(); });
            var small = function (id, icon, cap, fn) { me.button($p, id, { iconCls: 'toolbar__icon ' + icon, caption: cap }, true).on('click', function () { var a = api(); a && fn(a); me.done(); }); };
            small('tl-distrows', 'btn-elx-distrows', s.txtDistRows, function (a) { a.asc_DistributeTableCells(false); });
            small('tl-distcols', 'btn-elx-distcols', s.txtDistCols, function (a) { a.asc_DistributeTableCells(true); });
            var valign = function (v) { return function (a) { var p = new Asc.CTableProp(); p.put_CellsVAlign(v); a.tblApply(p); }; };
            small('tl-top', 'btn-elx-valign-top', s.txtAlignTop, valign(Asc.c_oAscVertAlignJc.Top));
            small('tl-middle', 'btn-elx-valign-middle', s.txtAlignMiddle, valign(Asc.c_oAscVertAlignJc.Center));
            small('tl-bottom', 'btn-elx-valign-bottom', s.txtAlignBottom, valign(Asc.c_oAscVertAlignJc.Bottom));

            // Cell Size: row height and column width of the selected cells
            this.tlH = this.spinner($p, 'tl-height', function (mm) { me.applyTable(function (p) { p.put_RowHeight(mm); }); });
            this.tlW = this.spinner($p, 'tl-width', function (mm) { me.applyTable(function (p) { p.put_ColumnWidth(mm); }); });

            // Text Direction and Cell Margins
            var dir = function (cap, icon, d) { return { caption: cap, iconCls: 'menu__icon ' + icon, checkable: true, checkmark: false, toggleGroup: 'elx-tl-dir', direction: d }; };
            this.tlDirection = this.button($p, 'tl-direction', { iconCls: 'toolbar__icon btn-elx-text-direction', caption: s.txtTextDirection, menu: new Common.UI.Menu({
                cls: 'ppm-toolbar shifted-right', items: [
                    dir(s.txtDirHorizontal, 'btn-text-orient-hor', Asc.c_oAscCellTextDirection.LRTB),
                    dir(s.txtDirRotateDown, 'btn-text-orient-rdown', Asc.c_oAscCellTextDirection.TBRL),
                    dir(s.txtDirRotateUp, 'btn-text-orient-rup', Asc.c_oAscCellTextDirection.BTLR)] }) }, true);
            this.tlDirection.menu.on('item:click', function (m, it) { me.applyTable(function (p) { p.put_CellsTextDirection(it.options.direction); }); });
            this.tlMargins = this.button($p, 'tl-margins', { iconCls: 'toolbar__icon btn-elx-cell-margins', caption: s.txtCellMargins, menu: new Common.UI.Menu({
                items: [{ caption: s.txtMarginsNormal, value: 'normal' }, { caption: s.txtMarginsNone, value: 'none' }, { caption: s.txtMarginsNarrow, value: 'narrow' },
                        { caption: s.txtMarginsModerate, value: 'moderate' }, { caption: s.txtMarginsWide, value: 'wide' }, { caption: '--' },
                        { caption: s.txtMarginsCustom, value: 'custom' }] }) }, true);
            this.tlMargins.menu.on('item:click', function (m, it) {
                if (it.value === 'custom') me.tableProperties();
                else me.applyCellMargins(it.value);
            });
        },

        // Chart Design: Chart Layouts | Chart Styles | Type | Data | Format pane
        build_chartdesign: function ($p) {
            var s = this.strings, me = this, _set = Common.enumLock;
            this.layout($p, [
                ['chartlayouts', this.big('cd-element') + this.big('cd-layout')],
                ['chartstyles', this.big('cd-colors') + this.big('cd-styles')],
                ['type', this.big('cd-type')],
                ['data', this.big('cd-data')],
                ['', this.big('cd-pane')]
            ]);
            $p.find('.separator').eq(3).addClass('invisible');
            // Common.UI.ChartRibbon, shared with the document and spreadsheet editors
            var dh = function () { return PE.getController('DocumentHolder'); };
            this.chartRibbon = Common.UI.ChartRibbon.create({
                strings: s,
                ids: {element: 'cd-element', layout: 'cd-layout', colors: 'cd-colors', styles: 'cd-styles'},
                lock: [_set.lostConnect, _set.noSlides, _set.disableOnStart],
                add: function (id, cmp) { cmp.render($p.find('#' + id)); me.controls.push(cmp); },
                api: function () { return me.api; },
                chart: function () { return me.api && me.api.asc_getChartSettings(true); },
                dh: dh,
                sourceMenu: function () { var h = dh(); return h && h.documentHolder && h.documentHolder.menuChartElement && h.documentHolder.menuChartElement.menu; },
                applyStyle: function (props, style) {
                    var p = new Asc.CAscChartProp();
                    props.putStyle(style);
                    p.put_ChartProperties(props);
                    me.api.ChartApply(p);
                },
                done: function () { me.done(); }
            });
            this.button($p, 'cd-type', { iconCls: 'toolbar__icon btn-elx-chart-type', caption: s.txtChangeChartType }).on('click', function () { me.chartCommand('type'); });
            this.button($p, 'cd-data', { iconCls: 'toolbar__icon btn-elx-chart-data', caption: s.txtEditData }).on('click', function () { me.chartCommand('data'); });
            this.button($p, 'cd-pane', { iconCls: 'toolbar__icon btn-elx-pane', caption: s.txtFormatPane }).on('click', function () { me.rightPanel(Asc.c_oAscTypeSelectElement.Chart); });
        },

        done: function () { Common.NotificationCenter.trigger('edit:complete', this.tb); },

        // ---- context ------------------------------------------------------------------------
        onFocusObject: function (objs) {
            var T = Asc.c_oAscTypeSelectElement, ctx = {};
            _.each(objs || [], function (o) {
                var t = o.get_ObjectType(), v = o.get_ObjectValue();
                if (t === T.Image) ctx.picture = v;
                else if (t === T.Shape) ctx.shape = v;
                else if (t === T.Table) ctx.table = v;
                else if (t === T.Chart) ctx.chart = v;
            });
            if (ctx.chart) { ctx.shape = null; ctx.picture = null; }
            if (ctx.picture) ctx.shape = null;
            if (ctx.table) ctx.shape = null;
            this._ctx = ctx;
            var vis = { shapeformat: !!ctx.shape, pictureformat: !!ctx.picture, tabledesign: !!ctx.table, tablelayout: !!ctx.table, chartdesign: !!ctx.chart };
            var me = this, active = this.activeTab();
            _.each(vis, function (v, tab) { me.setTabVisible(tab, v); });
            this._noApply = true;
            if (ctx.shape) this.updateSize(this.sfH, this.sfW, ctx.shape);
            if (ctx.picture) this.updateSize(this.pfH, this.pfW, ctx.picture);
            if (ctx.table) this.updateTable(ctx.table);
            this._noApply = false;
            if (active && vis[active] === false) this.tb.setTab(this._state.lastTab || 'home');
            var p = this._state.pending;
            if (p) {
                if (Date.now() > p.until) this._state.pending = null;
                else if (vis[p.tab]) { this._state.pending = null; this.tb.isCompact() || this.tb.setTab(p.tab); }
            }
        },

        updateSize: function (h, w, v) {
            h && v.get_Height && h.setValue(Common.Utils.Metric.fnRecalcFromMM(v.get_Height()).toFixed(2), true);
            w && v.get_Width && w.setValue(Common.Utils.Metric.fnRecalcFromMM(v.get_Width()).toFixed(2), true);
            var stroke = v.get_stroke && v.get_stroke();
            this._ctx.strokeWidth = stroke && stroke.get_width ? stroke.get_width() : 0;
        },

        updateTable: function (t) {
            var h = t.get_RowHeight && t.get_RowHeight(), w = t.get_ColumnWidth && t.get_ColumnWidth();
            this.tlH && this.tlH.setValue(h !== null && h !== undefined ? Common.Utils.Metric.fnRecalcFromMM(h).toFixed(2) : '', true);
            this.tlW && this.tlW.setValue(w !== null && w !== undefined ? Common.Utils.Metric.fnRecalcFromMM(w).toFixed(2) : '', true);
            var dir = t.get_CellsTextDirection && t.get_CellsTextDirection();
            this.tlDirection && _.each(this.tlDirection.menu.items, function (it) { it.setChecked && it.options.direction !== undefined && it.setChecked(it.options.direction === dir, true); });
            var look = t.get_TableLook && t.get_TableLook();
            if (!look) return;
            var vals = [look.get_FirstRow(), look.get_LastRow(), look.get_BandHor(), look.get_FirstCol(), look.get_LastCol(), look.get_BandVer()];
            _.each(this.tdChecks, function (c, i) { c.setValue(!!vals[i], true); });
        },

        updateMetric: function () {
            _.each(this.spinners || [], function (s) {
                s.setDefaultUnit(Common.Utils.Metric.getCurrentMetricName());
                s.setStep(Common.Utils.Metric.getCurrentMetric() === Common.Utils.Metric.c_MetricUnits.pt ? 1 : 0.1);
            });
        },

        updateColors: function () {
            var colors = Common.Utils.ThemeColor.getEffectColors();
            if (!colors) return;
            _.each(this.colorButtons || [], function (b) {
                b.picker && b.picker.updateColors(colors, Common.Utils.ThemeColor.getStandartColors());
                if (!b.currentColor) { b.currentColor = colors[4] && colors[4].color || colors[4]; b.setColor(b.currentColor); }
            });
            var std = Common.Utils.ThemeColor.getStandartColors();
            this.wa && this.wa.updateColors(colors, std);
            this.tdPen && this.tdPen.getPicker().updateColors(colors, std);
        },

        // inserting a table, picture, shape or chart opens its tab, whatever surface inserted it
        wrapInsertCommands: function (api) {
            if (api._elxCtxWrapped) return;
            api._elxCtxWrapped = true;
            var me = this;
            var wrap = function (name, tab) {
                var fn = api[name];
                if (typeof fn !== 'function') return;
                api[name] = function () {
                    me._state.pending = { tab: tab, until: Date.now() + 15000 };
                    return fn.apply(this, arguments);
                };
            };
            wrap('put_Table', 'tabledesign');
            wrap('AddImageUrl', 'pictureformat');
            wrap('asc_addImage', 'pictureformat');
            wrap('ChangeImageFromFile', 'pictureformat');
            wrap('StartAddShape', 'shapeformat');
            wrap('asc_addChartDrawingObject', 'chartdesign');
            wrap('AddTextArt', 'shapeformat');
        },

        // ---- commands (same API calls as the right panel) ----------------------------------
        rgb: function (c) { return Common.Utils.ThemeColor.getRgbColor(c); },

        applyShapeFill: function (color) {
            if (!this.api || !color) return;
            var props = new Asc.asc_CShapeProperty(), fill = new Asc.asc_CShapeFill();
            if (color === 'transparent') { fill.put_type(Asc.c_oAscFill.FILL_TYPE_NOFILL); fill.put_fill(null); }
            else { fill.put_type(Asc.c_oAscFill.FILL_TYPE_SOLID); fill.put_fill(new Asc.asc_CFillSolid()); fill.get_fill().put_color(this.rgb(color)); }
            props.put_fill(fill);
            this.api.ShapeApply(props);
            this.done();
        },

        applyShapeStroke: function (color) {
            if (!this.api || !color) return;
            var props = new Asc.asc_CShapeProperty(), stroke = new Asc.asc_CStroke();
            if (color === 'transparent') stroke.put_type(Asc.c_oAscStrokeType.STROKE_NONE);
            else {
                stroke.put_type(Asc.c_oAscStrokeType.STROKE_COLOR);
                stroke.put_color(this.rgb(color));
                var w = this._ctx && this._ctx.strokeWidth;
                stroke.put_width(w && w > 0 ? w : 0.75 * 25.4 / 72);
            }
            props.put_stroke(stroke);
            this._ctx && this._ctx.picture ? this.imageStroke(stroke) : this.api.ShapeApply(props);
            this.done();
        },

        imageStroke: function (stroke) {
            // a picture's border is its shape line
            var props = new Asc.asc_CShapeProperty();
            props.put_stroke(stroke);
            this.api.ShapeApply(props);
        },

        applySize: function (kind, w, h) {
            if (!this.api) return;
            var c = this._ctx || {}, cur = kind === 'image' ? c.picture : c.shape;
            if (!cur) return;
            var ratio = cur.get_Width() && cur.get_Height() ? cur.get_Width() / cur.get_Height() : 1;
            var lock = kind === 'image' || (cur.get_lockAspect && cur.get_lockAspect());
            if (w !== null && h === null) h = lock ? w / ratio : cur.get_Height();
            if (h !== null && w === null) w = lock ? h * ratio : cur.get_Width();
            var props = kind === 'image' ? new Asc.asc_CImgProperty() : new Asc.asc_CShapeProperty();
            props.put_Width(w);
            props.put_Height(h);
            kind === 'image' ? this.api.ImgApply(props) : this.api.ShapeApply(props);
        },

        changePicture: function (how) {
            var me = this;
            if (how === 1) {
                (new Common.Views.ImageFromUrlDialog({
                    handler: function (result, value) {
                        var url = result === 'ok' && value ? value.replace(/ /g, '') : '';
                        if (url) { var p = new Asc.asc_CImgProperty(); p.put_ImageUrl(url); me.api.ImgApply(p); }
                        me.done();
                    }
                })).show();
            } else if (how === 2) Common.NotificationCenter.trigger('storage:image-load', 'change');
            else { this.api.ChangeImageFromFile(); this.done(); }
        },

        resetPicture: function () {
            if (!this.api) return;
            var size = this.api.asc_getCropOriginalImageSize(), p = new Asc.asc_CImgProperty();
            p.put_ResetCrop(true);
            p.put_Width(size.get_ImageWidth());
            p.put_Height(size.get_ImageHeight());
            p.put_Rot(0);
            this.api.ImgApply(p);
            this.done();
        },

        crop: function (mode) {
            if (!this.api) return;
            if (mode === 0) { this.api.asc_startEditCrop(); this.pfCrop.toggle(true, true); }
            else if (mode === -1) { this.api.asc_endEditCrop(); this.pfCrop.toggle(false, true); }
            else if (mode === 1) this.api.asc_cropFill();
            else if (mode === 2) this.api.asc_cropFit();
            this.done();
        },

        applyTableLook: function (idx, on) {
            var c = this._ctx && this._ctx.table;
            if (!this.api || !c) return;
            var props = new Asc.CTableProp(), look = c.get_TableLook ? c.get_TableLook() : new Asc.CTablePropLook();
            ['put_FirstRow', 'put_LastRow', 'put_BandHor', 'put_FirstCol', 'put_LastCol', 'put_BandVer'][idx] &&
                look[['put_FirstRow', 'put_LastRow', 'put_BandHor', 'put_FirstCol', 'put_LastCol', 'put_BandVer'][idx]](on);
            props.put_TableLook(look);
            this.api.tblApply(props);
            this.done();
        },

        applyCellShading: function (color) {
            if (!this.api || !color) return;
            var props = new Asc.CTableProp(), bg = new Asc.CBackground();
            props.put_CellsBackground(bg);
            if (color === 'transparent') bg.put_Value(1);
            else { bg.put_Value(0); bg.put_Color(this.rgb(color)); }
            props.put_CellSelect(true);
            this.api.tblApply(props);
            this.done();
        },

        applyTable: function (fill) {
            if (!this.api) return;
            var props = new Asc.CTableProp();
            fill(props);
            this.api.tblApply(props);
            this.done();
        },

        // margins in mm (PowerPoint's presets: left, top, right, bottom); Normal gives the cells back the table default margins
        applyCellMargins: function (kind) {
            var v = CELL_MARGINS[kind];
            this.applyTable(function (p) {
                var cm = new Asc.CMargins();
                if (!v) cm.put_Flag(0);
                else {
                    cm.put_Left(v[0]); cm.put_Top(v[1]); cm.put_Right(v[2]); cm.put_Bottom(v[3]);
                    cm.put_Flag(2);
                }
                p.put_CellMargins(cm);
                p.put_CellSelect(true);
            });
        },

        // border presets of the Borders menu: the same CBorders as Word's Table Design tab
        applyCellBorders: function (border) {
            var me = this, borders = new Asc.CBorders(), visible = border !== '', size = this._borderSize, pen = this.tdPen;
            var set = function (putter, key) {
                if (border.indexOf(key) > -1 || !visible) {
                    var b = new Asc.asc_CTextBorder();
                    if (visible && size > 0) {
                        b.put_Value(1);
                        b.put_Size(parseFloat(size) * 25.4 / 72.0);
                        var color;
                        if (pen.isAutoColor()) { color = new Asc.asc_CColor(); color.put_auto(true); }
                        else color = this.rgb(pen.color);
                        b.put_Color(color);
                    } else b.put_Value(0);
                    borders[putter](b);
                }
            };
            set = _.bind(set, this);
            set('put_Left', 'l'); set('put_Top', 't'); set('put_Right', 'r'); set('put_Bottom', 'b'); set('put_InsideV', 'c'); set('put_InsideH', 'm');
            this.applyTable(function (p) { p.put_CellBorders(borders); p.put_CellSelect(true); });
        },

        // Table Properties (Placement, Margins, Alternative Text): the dialog of the right panel
        tableProperties: function () {
            var me = this, api = this.api, el = null;
            _.each(api ? api.getSelectedElements() : [], function (e) { if (e.get_ObjectType() === Asc.c_oAscTypeSelectElement.Table) el = e; });
            if (!el) return;
            (new PE.Views.TableSettingsAdvanced({
                tableProps: el.get_ObjectValue(),
                slideSize: PE.getController('Toolbar').currentPageSize,
                handler: function (result, value) {
                    result === 'ok' && me.api && me.api.tblApply(value.tableProps);
                    me.done();
                }
            })).show();
        },

        splitCells: function () {
            var me = this;
            (new Common.Views.InsertTableDialog({
                split: true,
                handler: function (result, value) {
                    result === 'ok' && me.api && me.api.SplitCell(value.columns, value.rows);
                    me.done();
                }
            })).show();
        },

        rightPanel: function (type) {
            var el = null;
            _.each(this.api ? this.api.getSelectedElements() : [], function (e) { if (e.get_ObjectType() === type) el = e; });
            var rm = this.getApplication().getController('RightMenu');
            if (!rm || !el) return null;
            rm.rightmenu && rm.rightmenu.isVisible && !rm.rightmenu.isVisible() && Common.NotificationCenter.trigger('rightmenu:show');
            rm.onDoubleClickOnObject(el);
            return rm;
        },

        chartCommand: function (what) {
            var rm = this.rightPanel(Asc.c_oAscTypeSelectElement.Chart),
                cs = rm && rm.rightmenu && rm.rightmenu.chartSettings;
            if (!cs) return;
            setTimeout(function () {
                var b = what === 'data' ? cs.btnEditData : cs.btnChartType;
                if (!b || b.isDisabled()) return;
                what === 'data' ? b.trigger('click', b) : b.cmpEl.find('button').first().click();
            }, 50);
        },

        hideAll: function () {
            var me = this;
            _.each(TABS, function (t) { me.setTabVisible(t, false); });
        },

        activeTab: function () {
            var $a = this.tb.$tabs.filter('.active').find('> a');
            return $a.length ? $a.data('tab') : null;
        },

        setTabVisible: function (tab, visible) {
            if (this._state.visible[tab] === visible) return;
            this._state.visible[tab] = visible;
            this.tb.setVisible(tab, visible);
            visible && tab === this.activeTab() && this.tb.processPanelVisible && this.tb.processPanelVisible(null, true);
        }
    });
});

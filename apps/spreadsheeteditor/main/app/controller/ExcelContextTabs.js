/*
 * SPDX-FileCopyrightText: 2026 ElixirOffice contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 *
 * SSE.Controllers.ExcelContextTabs: Excel's contextual tabs for the objects
 * of a sheet, shown at the right of the tab row while the object is selected:
 * - Chart Design (Change Chart Type, Select Data, Switch Row/Column) and
 *   Format (Format pane, Arrange) for a chart;
 * - Shape Format (Insert Shapes, Text Box, Format pane, Arrange) for a shape;
 * - Picture Format (Format pane, Arrange) for a picture.
 * The commands are those of the right panel and of Page Layout > Arrange
 * (twins, Common.UI.RibbonMirror): same handlers, locks and undo.
 * Inserting a chart opens Chart Design, like Excel.
 */

define([
    'core',
    'common/main/lib/component/Button',
    'common/main/lib/component/RibbonMirror'
], function () {
    'use strict';

    SSE.Controllers.ExcelContextTabs = Backbone.Controller.extend(_.extend({
        initialize: function () {
            Common.NotificationCenter.on('app:ready', _.bind(this.onAppReady, this));
        },

        ctrl: function (n) { return this.getApplication().getController(n); },

        onAppReady: function (mode) {
            this.mode = mode || {};
            if (!this.mode.isEdit || this.mode.isEditDiagram || this.mode.isEditMailMerge || this.mode.isEditOle) return;
            this.api = this.ctrl('Viewport').getApi();
            this.tb = this.ctrl('Toolbar').toolbar;
            if (!this.api || !this.tb || this.ready) return;
            this.ready = true;
            this.state = {};
            this.createTabs();
            this.api.asc_registerCallback('asc_onSelectionChanged', _.bind(this.onSelectionChanged, this));
            // a new chart opens Chart Design, like Excel
            var me = this, api = this.api;
            if (!api._elxCtxWrapped) {
                api._elxCtxWrapped = true;
                _.each(['asc_addChartDrawingObject', 'asc_addChartSpace'], function (fn) {
                    var orig = api[fn];
                    orig && (api[fn] = function () {
                        me.openOnNext = 'chartdesign';
                        return orig.apply(this, arguments);
                    });
                });
            }
        },

        // ---- tabs ---------------------------------------------------------------------
        panel: function (groups) {
            var html = '<section class="panel" role="tabpanel">';
            _.each(groups, function (g, i) {
                i && (html += '<div class="separator long"></div>');
                html += '<div class="group" data-group="' + g[0] + '">';
                _.each(g[1], function (id) { html += '<span class="btn-slot text x-huge" id="' + id + '"></span>'; });
                html += '</div>';
            });
            return $(html + '</section>');
        },

        bigBtn: function ($slot, cfg) {
            var b = new Common.UI.Button(_.extend({
                cls: 'btn-toolbar x-huge icon-top', dataHint: '1', dataHintDirection: 'bottom', dataHintOffset: 'small'
            }, cfg));
            $slot.length && b.render($slot);
            return b;
        },

        twin: function ($slot, src, caption) {
            if (!src || !$slot.length) return null;
            var t = Common.UI.RibbonMirror(src, caption ? {caption: caption} : {});
            t.render($slot);
            t.mirrorAttach();
            return t;
        },

        arrange: function (prefix) {
            return ['arrange', ['slot-elx-' + prefix + '-forward', 'slot-elx-' + prefix + '-backward', 'slot-elx-' + prefix + '-align', 'slot-elx-' + prefix + '-group']];
        },

        renderArrange: function ($p, prefix) {
            var tb = this.tb;
            this.twin($p.find('#slot-elx-' + prefix + '-forward'), tb.btnImgForward);
            this.twin($p.find('#slot-elx-' + prefix + '-backward'), tb.btnImgBackward);
            this.twin($p.find('#slot-elx-' + prefix + '-align'), tb.btnImgAlign);
            this.twin($p.find('#slot-elx-' + prefix + '-group'), tb.btnImgGroup);
        },

        createTabs: function () {
            var me = this, tb = this.tb, Doc = Common.Utils.documentSettingsType;
            var add = function (action, caption, $p) {
                $p.attr('data-tab', action);
                tb.addTab({action: action, caption: caption, aux: true, extcls: 'canedit'}, $p);
                tb.setVisible(action, false);
                tb.$el.find('.ribtab > a[data-tab=' + action + ']').closest('.ribtab').addClass('contextual');
            };
            // Chart Design
            var $cd = this.panel([['type', ['slot-elx-chart-type']], ['data', ['slot-elx-chart-switch', 'slot-elx-chart-data']]]);
            add('chartdesign', this.txtChartDesign, $cd);
            this.btnChartType = this.bigBtn($cd.find('#slot-elx-chart-type'), {caption: this.txtChangeChartType, iconCls: 'toolbar__icon btn-big-elx-chart-type'});
            this.btnChartSwitch = this.bigBtn($cd.find('#slot-elx-chart-switch'), {caption: this.txtSwitchRowCol, iconCls: 'toolbar__icon btn-big-elx-switch-rowcol'});
            this.btnChartData = this.bigBtn($cd.find('#slot-elx-chart-data'), {caption: this.txtSelectData, iconCls: 'toolbar__icon btn-big-elx-select-data'});
            this.btnChartType.on('click', function () { me.chartChangeType(); });
            this.btnChartSwitch.on('click', function () { me.chartSwitch(); });
            this.btnChartData.on('click', function () { me.chartSelectData(); });
            // Format (chart)
            var $cf = this.panel([['currentselection', ['slot-elx-chart-pane']], this.arrange('cf')]);
            add('chartformat', this.txtFormat, $cf);
            this.bigBtn($cf.find('#slot-elx-chart-pane'), {caption: this.txtFormatPane, iconCls: 'toolbar__icon btn-big-elx-format-pane'})
                .on('click', function () { me.openPane(Doc.Chart); });
            this.renderArrange($cf, 'cf');
            // Shape Format
            var $sf = this.panel([['insertshapes', ['slot-elx-shape-shapes', 'slot-elx-shape-text']], ['shapestyles', ['slot-elx-shape-pane']], this.arrange('sf')]);
            add('shapeformat', this.txtShapeFormat, $sf);
            this.twin($sf.find('#slot-elx-shape-shapes'), tb.btnInsertShape);
            this.twin($sf.find('#slot-elx-shape-text'), tb.btnInsertText);
            this.bigBtn($sf.find('#slot-elx-shape-pane'), {caption: this.txtFormatPane, iconCls: 'toolbar__icon btn-big-elx-format-pane'})
                .on('click', function () { me.openPane(Doc.Shape); });
            this.renderArrange($sf, 'sf');
            // Picture Format
            var $pf = this.panel([['picturestyles', ['slot-elx-pic-pane']], this.arrange('pf')]);
            add('pictureformat', this.txtPictureFormat, $pf);
            this.bigBtn($pf.find('#slot-elx-pic-pane'), {caption: this.txtFormatPane, iconCls: 'toolbar__icon btn-big-elx-format-pane'})
                .on('click', function () { me.openPane(Doc.Image); });
            this.renderArrange($pf, 'pf');
            var rt = this.ctrl('RibbonTabs');
            rt && rt.view && _.each(['chartdesign', 'chartformat', 'shapeformat', 'pictureformat'], function (t) {
                var $panel = tb.$el.find('.box-panels > .panel[data-tab=' + t + ']');
                tb.tagPanelGroups && tb.tagPanelGroups($panel);
            });
        },

        // ---- visibility -------------------------------------------------------------------
        onSelectionChanged: function (info) {
            if (!info || !this.tb) return;
            var T = Asc.c_oAscSelectionType, type = info.asc_getSelectionType(), tb = this.tb;
            var want = {
                chartdesign: type === T.RangeChart || type === T.RangeChartText,
                chartformat: type === T.RangeChart || type === T.RangeChartText,
                shapeformat: type === T.RangeShape || type === T.RangeShapeText,
                pictureformat: type === T.RangeImage
            };
            var me = this, changed = false;
            _.each(want, function (on, tab) {
                if (me.state[tab] === on) return;
                me.state[tab] = on;
                changed = true;
                if (!on && tb.isTabActive(tab)) tb.setTab('home');
                tb.setVisible(tab, on);
            });
            if (this.openOnNext && want[this.openOnNext]) {
                tb.setTab(this.openOnNext);
                this.openOnNext = null;
            }
            changed && tb.processPanelVisible && tb.processPanelVisible(null, true);
            var editable = !info.asc_getLocked || !info.asc_getLocked();
            this.btnChartType && this.btnChartType.setDisabled(!editable);
            this.btnChartData && this.btnChartData.setDisabled(!editable);
            this.btnChartSwitch && this.btnChartSwitch.setDisabled(!editable);
        },

        openPane: function (type) {
            var rm = this.ctrl('RightMenu');
            rm && rm.onRightMenuOpen && rm.onRightMenuOpen(type);
        },

        // ---- chart commands (those of the right panel, SSE.Views.ChartSettings) -----------
        chartType: function () {
            var objs = this.api.asc_getGraphicObjectProps() || [], type;
            _.each(objs, function (o) {
                var v = o.asc_getObjectValue && o.asc_getObjectValue(), cp = v && v.asc_getChartProperties && v.asc_getChartProperties();
                if (cp && type === undefined) type = cp.getType();
            });
            return type;
        },

        chartChangeType: function () {
            SSE.Views.ChartSettings.prototype.onChangeType.call({api: this.api, toolbar: this.tb, _state: {ChartType: this.chartType()}});
        },

        chartSelectData: function () {
            SSE.Views.ChartSettings.prototype.onSelectData.call({api: this.api});
        },

        chartSwitch: function () {
            SSE.Views.ChartSettings.prototype.onSwitch.call({api: this.api, errorMaxRows: SSE.Views.ChartSettings.prototype.errorMaxRows});
            Common.NotificationCenter.trigger('edit:complete', this.tb);
        },

        txtChartDesign: 'Chart Design',
        txtFormat: 'Format',
        txtShapeFormat: 'Shape Format',
        txtPictureFormat: 'Picture Format',
        txtChangeChartType: 'Change Chart Type',
        txtSwitchRowCol: 'Switch Row/Column',
        txtSelectData: 'Select Data',
        txtFormatPane: 'Format Pane'
    }, SSE.Controllers.ExcelContextTabs || {}));
});

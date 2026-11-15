/*
 * SPDX-FileCopyrightText: 2026 ElixirOffice contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 *
 * Common.UI.ChartRibbon: the Chart Design controls shared by the Word, Excel and PowerPoint
 * contextual tabs, like the Chart Layouts and Chart Styles groups of Office:
 *   Add Chart Element  axes, axis titles, chart title, data labels, error bars, gridlines, legend,
 *                      trendline, up/down bars. The menu is a copy of the chart element menu of the
 *                      document holder (the one of the chart's floating button), and runs the same
 *                      handlers (onChartElement, updateChartElementMenu), so the engine calls
 *                      (asc_ChartSettings.setDisplay*) are the existing ones.
 *   Quick Layout       a gallery of element combinations, applied with the same setDisplay* calls.
 *   Change Colors      Colorful (the engine's palette) and Monochromatic 1 to 6 (shades of each theme accent),
 *                      asc_ChartSettings.setChartColors / getChartColorsAccent.
 *   Chart Styles       a gallery of the chart style presets of the chart type (asc_getChartPreviews),
 *                      applied like the Style combo of the right panel. Optional: Word has its own.
 * The data table is not offered: the engine stores it but does not draw it yet.
 *
 * The editors supply a small host object:
 *   host.add(slotId, component)  render the component into its slot of the tab, register it for locks
 *   host.lock                    lock list of the new controls
 *   host.strings                 txtAddChartElement, txtQuickLayout, txtChangeColors, txtChartStyles, txtLayout,
 *                                txtColorful, txtMonochromatic (followed by the accent number)
 *   host.ids                     {element, layout, colors, styles}: slot ids (styles: omit for no gallery)
 *   host.api()                   the editor api
 *   host.chart()                 the asc_ChartSettings of the selected chart (fresh), or null
 *   host.dh()                    the DocumentHolder controller (onChartElement, updateChartElementMenu)
 *   host.sourceMenu()            the chart element menu of the document holder view
 *   host.applyStyle(props, id)   applies a chart style (the editor's own ImgApply/ChartApply call)
 *   host.transaction             true when the history supports transactions (spreadsheet): one undo step
 *   host.done()                  give the focus back to the document
 */
define([
    'core',
    'common/main/lib/component/Button',
    'common/main/lib/component/DataView',
    'common/main/lib/component/Menu'
], function () {
    'use strict';

    var uid = 0;

    // Quick Layout presets: title, legend position, data labels, major gridlines of the value axis, axis titles
    var LAYOUTS = [
        {title: true, legend: 'right', grid: true},
        {title: true, legend: 'top', labels: true},
        {title: true, legend: 'bottom', grid: true},
        {legend: 'bottom', labels: true, grid: true},
        {title: true, legend: 'right', axisTitles: true, grid: true},
        {title: true, labels: true},
        {grid: true},
        {title: true, legend: 'left', labels: true, grid: true}
    ];

    // Change Colors: -1 is the colorful palette, 0 to 5 the monochromatic palette of accent 1 to 6
    var ACCENTS = [0, 1, 2, 3, 4, 5];

    // a thumbnail of a layout: title, legend, labels, gridlines and axis titles as in the layout
    function thumbnail(def) {
        var x0 = 7, x1 = 43, y0 = def.title ? 11 : 5, y1 = def.axisTitles ? 26 : 30, g = '';
        if (def.legend === 'right') x1 = 33;
        if (def.legend === 'left') x0 = 17;
        if (def.legend === 'top') y0 += 5;
        if (def.legend === 'bottom') y1 -= 4;
        var h = y1 - y0, i;
        if (def.grid) for (i = 1; i <= 3; i++) g += '<line x1="' + x0 + '" x2="' + x1 + '" y1="' + (y1 - h * i / 3).toFixed(1) + '" y2="' + (y1 - h * i / 3).toFixed(1) + '" opacity=".3"/>';
        g += '<line x1="' + x0 + '" x2="' + x1 + '" y1="' + y1 + '" y2="' + y1 + '"/>';
        var bw = (x1 - x0 - 8) / 3;
        for (i = 0; i < 3; i++) {
            var bh = h * [0.5, 0.85, 0.65][i], bx = x0 + 2 + i * (bw + 2);
            g += '<rect x="' + bx.toFixed(1) + '" y="' + (y1 - bh).toFixed(1) + '" width="' + bw.toFixed(1) + '" height="' + bh.toFixed(1) + '" fill="currentColor" stroke="none" opacity=".55"/>';
            if (def.labels) g += '<line x1="' + (bx + 1).toFixed(1) + '" x2="' + (bx + bw - 1).toFixed(1) + '" y1="' + (y1 - bh - 2).toFixed(1) + '" y2="' + (y1 - bh - 2).toFixed(1) + '" stroke-width="1.5"/>';
        }
        if (def.title) g += '<line x1="15" x2="35" y1="5" y2="5" stroke-width="2"/>';
        if (def.axisTitles) g += '<line x1="17" x2="33" y1="33" y2="33" opacity=".7"/><line x1="3" x2="3" y1="14" y2="24" opacity=".7"/>';
        if (def.legend) {
            var lx = {right: 37, left: 3}[def.legend], ly = {top: 9, bottom: 33}[def.legend];
            if (lx !== undefined) g += '<rect x="' + lx + '" y="14" width="3" height="3" fill="currentColor" stroke="none"/><line x1="' + (lx + 5) + '" x2="' + (lx + 9) + '" y1="15.5" y2="15.5"/><rect x="' + lx + '" y="20" width="3" height="3" fill="currentColor" stroke="none" opacity=".55"/><line x1="' + (lx + 5) + '" x2="' + (lx + 9) + '" y1="21.5" y2="21.5"/>';
            else g += '<rect x="16" y="' + (ly - 1.5) + '" width="3" height="3" fill="currentColor" stroke="none"/><line x1="21" x2="26" y1="' + ly + '" y2="' + ly + '"/><rect x="29" y="' + (ly - 1.5) + '" width="3" height="3" fill="currentColor" stroke="none" opacity=".55"/>';
        }
        return '<svg width="48" height="36" viewBox="0 0 48 36" fill="none" stroke="currentColor" stroke-width="1">' + g + '</svg>';
    }

    function labelPos(type) {
        var t = Asc.c_oAscChartTypeSettings, P = Asc.c_oAscChartDataLabelsPos;
        if (type === t.pie || type === t.pie3d) return P.bestFit;
        if (type === t.barNormal || type === t.hBarNormal) return P.outEnd;
        return P.ctr;
    }

    var LEGEND = function () {
        var L = Asc.c_oAscChartLegendShowSettings;
        return {right: L.right, left: L.left, top: L.top, bottom: L.bottom};
    };

    // the items of the document holder's chart element menu, as configs for a new menu
    function cloneItems(items) {
        return _.map(items, function (it) {
            var cfg = {caption: it.caption, value: it.value, checkable: it.checkable, toggleGroup: it.toggleGroup, stopPropagation: it.options && it.options.stopPropagation};
            if (it.menu && it.menu.items) cfg.menu = new Common.UI.Menu({cls: 'shifted-right', menuAlign: 'tl-tr', items: cloneItems(it.menu.items)});
            return cfg;
        });
    }

    Common.UI.ChartRibbon = {

        LAYOUTS: LAYOUTS,
        ACCENTS: ACCENTS,
        thumbnail: thumbnail,

        // What a preset would change: the steps of it that the chart does not already match (each step is an
        // undo step of its own where the history has no transactions, so only the needed ones are run)
        layoutSteps: function (p, def) {
            var T = Asc.c_oAscChartTitleShowSettings, L = Asc.c_oAscChartLegendShowSettings, type = p.getType();
            var hor = p.getHorAxesProps && p.getHorAxesProps() || [], ver = p.getVertAxesProps && p.getVertAxesProps() || [], axes = hor.concat(ver);
            var steps = {};
            steps.title = def.title ? p.getTitle() !== T.noOverlay : p.getTitle() !== T.none;
            steps.legend = (def.legend ? LEGEND()[def.legend] : L.none) !== p.getLegendPos();
            steps.labels = (def.labels ? labelPos(type) : Asc.c_oAscChartDataLabelsPos.none) !== p.getDataLabelsPos();
            if (axes.length) {
                var grids = _.map(axes, function (a) { return a.getGridlines(); });
                var major = _.filter(grids, function (g) { return g === 1; }).length, none = _.filter(grids, function (g) { return !g; }).length;
                steps.grid = def.grid ? !(major === 1 && none === grids.length - 1) : none !== grids.length;
                var titles = def.axisTitles ? 1 : 0;
                steps.axisTitles = !(hor[0] && hor[0].getLabel() === titles && ver[0] && ver[0].getLabel() === titles);
            }
            return steps;
        },

        // Quick Layout: sets the chart elements of a preset, one setDisplay* call each, on a fresh chart object
        applyLayout: function (host, def) {
            var p = host.chart();
            if (!p) return;
            var steps = this.layoutSteps(p, def), type = p.getType();
            var H = typeof AscCommon !== 'undefined' && AscCommon.History, tx = !!(host.transaction && H && H.StartTransaction);
            if (tx) {
                H.Create_NewPoint(AscDFH.historyitem_type_ChartSpace);
                H.StartTransaction();
            }
            try {
                var fresh = function () { return host.chart() || p; };
                steps.title && fresh().setDisplayChartTitle(!!def.title, false);
                steps.legend && fresh().setDisplayLegend(!!def.legend, def.legend ? LEGEND()[def.legend] : Asc.c_oAscChartLegendShowSettings.none);
                steps.labels && fresh().setDisplayDataLabels(!!def.labels, def.labels ? labelPos(type) : false);
                steps.grid && fresh().setDisplayGridlines(!!def.grid, false, false, false);
                steps.axisTitles && fresh().setDisplayAxisTitles(!!def.axisTitles, false, !!def.axisTitles, false, false);
            } finally {
                tx && H.EndTransaction();
            }
        },

        // Builds the controls; returns {controls: [...], ...buttons}
        create: function (host) {
            var s = host.strings, ids = host.ids, lock = host.lock, n = ++uid, me = this, previewHooked = false;
            var mk = function (id, cfg) {
                var b = new Common.UI.Button(_.extend({
                    cls: 'btn-toolbar x-huge icon-top', lock: lock, dataHint: '1', dataHintDirection: 'bottom', dataHintOffset: 'small'
                }, cfg));
                host.add(id, b);
                return b;
            };
            var holder = function (id, w) { return {template: _.template('<div id="' + id + '" style="width: ' + w + 'px; margin: 0 4px;"></div>')}; };
            var controls = [];

            // ---- Add Chart Element -------------------------------------------------------------
            var btnElement = mk(ids.element, {
                iconCls: 'toolbar__icon btn-elx-chart-add-element', caption: s.txtAddChartElement,
                menu: new Common.UI.Menu({items: [{caption: s.txtAddChartElement}]})
            });
            controls.push(btnElement);
            var elementsBuilt = false;
            btnElement.menu.on('show:before', function (menu) {
                var dh = host.dh(), props = host.chart(), source = host.sourceMenu();
                if (!dh || !props) return;
                if (!elementsBuilt && source && source.items && source.items.length) {
                    elementsBuilt = true;
                    menu.removeAll();
                    _.each(cloneItems(source.items), function (cfg) { menu.addItem(cfg); });
                    _.each(menu.items, function (top) {
                        top.menu && _.each(top.menu.items, function (sub) {
                            sub.on('click', function () {
                                var p = host.chart();
                                if (!p) return;
                                dh.chartProps = p;
                                dh.onChartElement(sub.menu, sub);
                            });
                        });
                    });
                }
                elementsBuilt && dh.updateChartElementMenu(menu, props);
            });
            btnElement.menu.on('hide:after', function () { host.done && host.done(); });

            // ---- Quick Layout ------------------------------------------------------------------
            var lid = 'elx-chart-layouts-' + n;
            var btnLayout = mk(ids.layout, {
                iconCls: 'toolbar__icon btn-elx-chart-quick-layout', caption: s.txtQuickLayout,
                menu: new Common.UI.Menu({style: 'min-width: 0;', items: [holder(lid, 232)]})
            });
            controls.push(btnLayout);
            var layoutBuilt = false;
            btnLayout.menu.on('show:before', function (menu) {
                if (layoutBuilt) return;
                layoutBuilt = true;
                var store = new Common.UI.DataViewStore(_.map(LAYOUTS, function (def, i) {
                    return {svg: thumbnail(def), tip: (s.txtLayout || 'Layout') + ' ' + (i + 1), def: def};
                }));
                var picker = new Common.UI.DataView({
                    el: $('#' + lid, menu.$el),
                    store: store,
                    parentMenu: menu,
                    outerMenu: {menu: menu, index: 0},
                    showLast: false,
                    restoreHeight: 300,
                    itemTemplate: _.template('<div class="item-chart-layout" title="<%= tip %>" style="width:52px;height:40px;padding:2px;"><%= svg %></div>')
                });
                picker.on('item:click', function (picker, item, record, e) {
                    record && me.applyLayout(host, record.get('def'));
                    host.done && host.done();
                    if (e.type !== 'click') menu.hide();
                });
                menu.setInnerMenu([{menu: picker, index: 0}]);
            });

            // ---- Change Colors -----------------------------------------------------------------
            var btnColors = mk(ids.colors, {
                iconCls: 'toolbar__icon btn-elx-chart-colors', caption: s.txtChangeColors,
                menu: new Common.UI.Menu({items: [{caption: s.txtColorful, value: -1, checkable: true, toggleGroup: 'elx-chart-palette-' + n}, {caption: '--'}].concat(_.map(ACCENTS, function (a) {
                    return {caption: s.txtMonochromatic + ' ' + (a + 1), value: a, checkable: true, toggleGroup: 'elx-chart-palette-' + n};
                }))})
            });
            controls.push(btnColors);
            btnColors.menu.on('show:before', function (menu) {
                var p = host.chart(), cur = p && p.getChartColorsAccent ? p.getChartColorsAccent() : null;
                _.each(menu.items, function (it) { it.value !== undefined && it.setChecked(it.value === cur, true); });
            });
            btnColors.menu.on('item:click', function (menu, item) {
                var p = host.chart();
                if (p && p.setChartColors) p.setChartColors(item.value);
                host.done && host.done();
            });

            // ---- Chart Styles gallery ----------------------------------------------------------
            var btnStyles;
            if (ids.styles) {
                var sid = 'elx-chart-styles-' + n, stylesStore = null;
                btnStyles = mk(ids.styles, {
                    iconCls: 'toolbar__icon btn-elx-chart-styles', caption: s.txtChartStyles,
                    menu: new Common.UI.Menu({style: 'min-width: 0;', items: [holder(sid, 330)]})
                });
                controls.push(btnStyles);
                var picker = null;
                btnStyles.menu.on('show:before', function (menu) {
                    var api = host.api(), props = host.chart();
                    if (!api || !props) return;
                    if (!previewHooked) {
                        previewHooked = true;
                        api.asc_registerCallback('asc_onAddChartStylesPreview', function (styles) {
                            stylesStore && _.each(styles || [], function (st) {
                                var rec = stylesStore.findWhere({data: st.asc_getName()});
                                rec && rec.set('imageUrl', st.asc_getImage());
                            });
                        });
                    }
                    var type = props.getType(), t = Asc.c_oAscChartTypeSettings;
                    var combo = type === t.comboBarLine || type === t.comboBarLineSecondary || type === t.comboAreaBar || type === t.comboCustom;
                    var list = !combo && type >= 0 ? (api.asc_getChartPreviews(type, undefined, true) || []) : [];
                    var data = _.map(list, function (st) { return {imageUrl: st.asc_getImage(), data: st.asc_getName(), tip: s.txtChartStyles + ' ' + st.asc_getName()}; });
                    if (!picker) {
                        stylesStore = new Common.UI.DataViewStore(data);
                        picker = new Common.UI.DataView({
                            el: $('#' + sid, menu.$el),
                            store: stylesStore,
                            parentMenu: menu,
                            outerMenu: {menu: menu, index: 0},
                            showLast: false,
                            restoreHeight: 300,
                            itemTemplate: _.template('<div class="item-chart-style" title="<%= tip %>" style="width:50px;height:50px;"><% if (imageUrl) { %><img src="<%= imageUrl %>" id="<%= id %>" style="width:50px;height:50px;"><% } %></div>')
                        });
                        picker.on('item:click', function (picker, item, record, e) {
                            var p = host.chart();
                            record && p && host.applyStyle(p, record.get('data'));
                            host.done && host.done();
                            if (e.type !== 'click') menu.hide();
                        });
                        menu.setInnerMenu([{menu: picker, index: 0}]);
                    } else {
                        stylesStore.reset(data);
                    }
                    var rec = stylesStore.findWhere({data: props.getStyle()});
                    rec ? picker.selectRecord(rec, true) : picker.deselectAll && picker.deselectAll(true);
                    data.length && api.asc_generateChartPreviews(type, _.pluck(data, 'data'));
                });
            }

            return {
                controls: controls, btnElement: btnElement, btnLayout: btnLayout, btnColors: btnColors, btnStyles: btnStyles
            };
        }
    };
});

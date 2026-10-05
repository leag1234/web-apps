/*
 * SPDX-FileCopyrightText: 2026 ElixirOffice contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 *
 * DE.Controllers.ContextTabs: Word-like contextual ribbon tabs.
 *
 * Tabs (appended at the right of the tab row and highlighted, like Word):
 *   tabledesign, tablelayout   cursor in a table
 *   pictureformat              a picture is selected
 *   shapeformat                a shape or text box is selected (or its text is edited)
 *   equation                   cursor in an equation
 *   headerfooter               (existing HeaderFooterTab) header or footer being edited
 * A tab appears with its context and goes away with it; when the active tab goes
 * away the ribbon returns to the tab that was active before. Inserting a table, a
 * picture, a shape or an equation opens the matching tab, like Word.
 *
 * The right settings panel stays available as a secondary surface.
 */

define([
    'core',
    'documenteditor/main/app/view/ContextTabs'
], function () {
    'use strict';

    var CONTEXT_TABS = ['headerfooter', 'tabledesign', 'tablelayout', 'pictureformat', 'shapeformat', 'equation'];
    var INSERT_TAB = {table: 'tabledesign', picture: 'pictureformat', shape: 'shapeformat', equation: 'equation'};

    DE.Controllers.ContextTabs = Backbone.Controller.extend(_.extend({
        models: [],
        collections: [],
        views: ['ContextTabs'],

        initialize: function () {
            var me = this;
            this._state = {visible: {}, pending: null, lastTab: 'home'};
            this.addListeners({
                'Toolbar': {
                    'render:after': function (toolbar) { me.onToolbarRender(toolbar); },
                    'tab:active': function (tab) { me.onTabActive(tab); },
                    'group:launcher': function (group) { me.onGroupLauncher(group); }
                }
            });
            Common.NotificationCenter.on('app:ready', this.onAppReady.bind(this));
            Common.NotificationCenter.on('settings:unitschanged', function () { me.view && me.view.updateMetricUnit(); });
        },

        onLaunch: function () {},

        appOptions: function () {
            var main = this.getApplication().getController('Main');
            return (main && main.appOptions) || {};
        },

        // ---- setup ---------------------------------------------------------------------
        onToolbarRender: function (toolbar) {
            var mode = this.appOptions();
            if (!mode.isEdit || mode.isEditMailMerge || mode.isEditDiagram || mode.isEditOle || this.view) return;
            this.toolbar = toolbar;
            this.toolbarCtrl = this.getApplication().getController('Toolbar');
            this.api = this.getApplication().getController('Viewport').getApi();
            this.view = this.createView('ContextTabs', {toolbar: toolbar, mode: mode});
            var panels = this.view.createPanels(), me = this, v = this.view;

            // group labels of the new panels (see Toolbar.updateGroupLabels)
            _.each(_.keys(v.constructor.prototype), function (k) {
                if (/^txtGroup/.test(k) && toolbar[k] === undefined) toolbar[k] = v[k];
            });

            var tabs = [
                {action: 'tabledesign', caption: v.txtTabTableDesign, dataHintTitle: 'JT'},
                {action: 'tablelayout', caption: v.txtTabTableLayout, dataHintTitle: 'JL'},
                {action: 'pictureformat', caption: v.txtTabPictureFormat, dataHintTitle: 'JP'},
                {action: 'shapeformat', caption: v.txtTabShapeFormat, dataHintTitle: 'JD'},
                {action: 'equation', caption: v.txtTabEquation, dataHintTitle: 'JE'}
            ];
            _.each(tabs, function (tab) {
                tab.extcls = 'canedit contextual';
                tab.layoutname = 'toolbar-' + tab.action;
                toolbar.addTab(tab, panels[tab.action], toolbar.getLastTabIdx());
            });
            Array.prototype.push.apply(toolbar.lockControls, v.getButtons());
            Common.Utils.lockControls(Common.enumLock.disableOnStart, true, {array: v.getButtons()});

            this.bindEvents();
            this.setApi(this.api);
            // the tab strip exists right after this event; Header & Footer is added later in onAppShowed
            setTimeout(function () { me.arrangeTabs(); }, 0);
        },

        arrangeTabs: function () {
            var tb = this.toolbar;
            if (!tb || !tb.$tabs) return;
            var $ul = tb.$tabs.first().parent();
            _.each(CONTEXT_TABS, function (tab) {
                var $li = $ul.find('> li > a[data-tab=' + tab + ']').parent();
                $li.addClass('contextual').css('display', 'none');
                $li.appendTo($ul);
            });
            tb.$tabs = $ul.find('> li');
            var $hf = $ul.find('> li > a[data-tab=headerfooter]');
            $hf.attr('data-hint-title', 'JH');
            this._state.visible = {};
        },

        onAppReady: function () {
            if (!this.view) return;
            var me = this;
            Common.Utils.lockControls(Common.enumLock.disableOnStart, false, {array: this.view.getButtons()});
            this.view.updateMetricUnit();
            this.updateColors();
            _.each(this.view.spinners, function (s) { s.on('inputleave', function () { Common.NotificationCenter.trigger('edit:complete', me.toolbar); }); });
            // follow the existing gridlines setting
            this.api && this.view.btnTblGridlines.toggle(!!this.api.get_ShowTableEmptyLine(), true);
            this.api && this.api.asc_GetMathInputType && this.updateEqInput();
        },

        setApi: function (api) {
            var me = this;
            if (!api) return;
            this.api = api;
            api.asc_registerCallback('asc_onFocusObject', _.bind(this.onApiFocusObject, this));
            api.asc_registerCallback('asc_onBeginTableStylesPreview', _.bind(this.onBeginTableStylesPreview, this));
            api.asc_registerCallback('asc_onAddTableStylesPreview', _.bind(this.onAddTableStylesPreview, this));
            api.asc_registerCallback('asc_onEndTableStylesPreview', _.bind(this.onEndTableStylesPreview, this));
            api.asc_registerCallback('asc_onTableDrawModeChanged', function (v) { me.view.btnTblDraw.toggle(!!v, true); });
            api.asc_registerCallback('asc_onTableEraseModeChanged', function (v) { me.view.btnTblEraser.toggle(!!v, true); });
            api.asc_registerCallback('asc_onChangeCropState', function (state) { me.view.btnPicCrop.toggle(!!state, true); });
            api.asc_registerCallback('asc_onSendThemeColors', function () { setTimeout(function () { me.updateColors(); }, 0); });
            api.asc_registerCallback('asc_onCoAuthoringDisconnect', function () { me.hideAll(); });
            this.wrapInsertCommands(api);
        },

        // Inserting an object opens its contextual tab (Word behaviour). The commands are
        // wrapped once, whatever surface calls them (ribbon, context menu, dialogs, AI).
        wrapInsertCommands: function (api) {
            var me = this;
            var wrap = function (name, kind, ttl, cond) {
                var orig = api[name];
                if (!_.isFunction(orig) || orig._elxWrapped) return;
                var fn = function () {
                    if (!cond || cond.apply(this, arguments))
                        me._state.pending = {kind: kind, until: Date.now() + (ttl || 3000)};
                    return orig.apply(this, arguments);
                };
                fn._elxWrapped = true;
                api[name] = fn;
            };
            wrap('put_Table', 'table');
            wrap('asc_ConvertTextToTable', 'table');
            wrap('asc_addImage', 'picture', 120000);
            wrap('AddImageUrl', 'picture', 60000);
            wrap('StartAddShape', 'shape', 120000, function (type, start) { return !!start; });
            wrap('AddTextArt', 'shape');
            wrap('asc_AddMath', 'equation');
            wrap('asc_AddMath2', 'equation');
        },

        // ---- context detection -----------------------------------------------------------
        onApiFocusObject: function (selected) {
            if (!this.view || !this.toolbar || !this.toolbar.$tabs) return;
            var ctx = {table: null, picture: null, shape: null, equation: false, para: null, locked: {}};
            for (var i = 0; i < selected.length; i++) {
                var type = selected[i].get_ObjectType(), value = selected[i].get_ObjectValue();
                if (type === Asc.c_oAscTypeSelectElement.Table) {
                    ctx.table = value;
                    ctx.locked.table = value.get_Locked();
                } else if (type === Asc.c_oAscTypeSelectElement.Image) {
                    if (value.get_ChartProperties()) continue;
                    var shp = value.get_ShapeProperties();
                    if (shp) {
                        if (shp.asc_getFromSmartArt && shp.asc_getFromSmartArt()) continue;
                        this._picShapeProps = shp;
                        ctx.shape = value;
                    } else ctx.picture = value;
                    ctx.locked.image = value.get_Locked();
                } else if (type === Asc.c_oAscTypeSelectElement.Math) {
                    ctx.equation = true;
                } else if (type === Asc.c_oAscTypeSelectElement.Paragraph) {
                    ctx.para = value;
                }
            }
            // a picture is also reported with shape properties (its border): it is a picture, not a shape
            if (ctx.picture) ctx.shape = null;
            this._ctx = ctx;
            var vis = {
                tabledesign: !!ctx.table, tablelayout: !!ctx.table,
                pictureformat: !!ctx.picture, shapeformat: !!ctx.shape, equation: !!ctx.equation
            };
            var me = this, active = this.activeTab();
            _.each(vis, function (v, tab) { me.setTabVisible(tab, v); });

            if (ctx.table) this.updateTable(ctx.table);
            if (ctx.picture) this.updatePicture(ctx.picture);
            if (ctx.shape) this.updateShape(ctx.shape);
            if (ctx.equation) this.updateEquation();

            // the active contextual tab went away: back to the tab used before it
            if (active && vis[active] === false) this.toolbar.setTab(this._state.lastTab || 'home');

            // an object was just inserted: open its tab
            var p = this._state.pending;
            if (p) {
                if (Date.now() > p.until) this._state.pending = null;
                else {
                    var hit = (p.kind === 'table' && ctx.table) || (p.kind === 'picture' && ctx.picture) ||
                              (p.kind === 'shape' && ctx.shape) || (p.kind === 'equation' && ctx.equation);
                    if (hit) {
                        this._state.pending = null;
                        if (!this.toolbar.isCompact()) this.toolbar.setTab(INSERT_TAB[p.kind]);
                    }
                }
            }
        },

        activeTab: function () {
            var $a = this.toolbar.$tabs.filter('.active').find('> a');
            return $a.length ? $a.data('tab') : null;
        },

        setTabVisible: function (tab, visible) {
            if (this._state.visible[tab] === visible) return;
            this._state.visible[tab] = visible;
            this.toolbar.setVisible(tab, visible);
            visible && this.toolbar.processPanelVisible && tab === this.activeTab() && this.toolbar.processPanelVisible(null, true);
        },

        hideAll: function () {
            var me = this;
            _.each(['tabledesign', 'tablelayout', 'pictureformat', 'shapeformat', 'equation'], function (t) { me.setTabVisible(t, false); });
        },

        onTabActive: function (tab) {
            if (CONTEXT_TABS.indexOf(tab) < 0) this._state.lastTab = tab;
            else if (tab === 'tabledesign') this.ensureTableStyles();
            if (tab === 'equation') this.ensureEquationStructures();
        },

        lockFor: function (tab, locked) {
            Common.Utils.lockControls('elx-ctx-locked', !!locked, {array: this.view.controls[tab]});
        },

        // ---- state of the controls -----------------------------------------------------
        updateTable: function (props) {
            var v = this.view, s = this._state;
            this.lockFor('tabledesign', props.get_Locked());
            this.lockFor('tablelayout', props.get_Locked());
            this._tableProps = props;
            var look = props.get_TableLook();
            if (look) {
                v.chHeader.setValue(!!look.get_FirstRow(), true);
                v.chTotal.setValue(!!look.get_LastRow(), true);
                v.chBanded.setValue(!!look.get_BandHor(), true);
                v.chFirst.setValue(!!look.get_FirstCol(), true);
                v.chLast.setValue(!!look.get_LastCol(), true);
                v.chColBanded.setValue(!!look.get_BandVer(), true);
            }
            var styleId = props.get_TableStyle();
            if (s.tableStyle !== styleId) {
                s.tableStyle = styleId;
                this.selectTableStyle();
            }
            var h = props.get_RowHeight(), w = props.get_ColumnWidth();
            v.spnTblHeight.setValue((h !== null && h !== undefined) ? Common.Utils.Metric.fnRecalcFromMM(h) : '', true);
            v.spnTblWidth.setValue((w !== null && w !== undefined) ? Common.Utils.Metric.fnRecalcFromMM(w) : '', true);
            var rows = props.get_RowsInHeader();
            v.btnTblRepeatHeader.toggle(!!rows, true);
            v.btnTblRepeatHeader.setDisabled(rows === null || props.get_Locked());
            if (this.api) {
                v.btnTblMerge.setDisabled(!this.api.CheckBeforeMergeCells() || props.get_Locked());
                v.btnTblSplit.setDisabled(!this.api.CheckBeforeSplitCells() || props.get_Locked());
            }
            var valign = props.get_CellsVAlign(), jc = this._ctx.para ? this._ctx.para.get_Jc() : null;
            _.each(v.btnsCellAlign, function (b) {
                b.toggle(b.options.valign === valign && b.options.halign === jc, true);
            });
            var dir = props.get_CellsTextDirection();
            _.each(v.btnTblDirection.menu.items, function (it) { it.setChecked(it.options.direction === dir, true); });
            var bg = props.get_CellsBackground();
            if (bg && bg.get_Value() == 0 && bg.get_Color()) {
                var c = bg.get_Color();
                v.btnShading.setColor(Common.Utils.ThemeColor.getHexColor(c.get_r(), c.get_g(), c.get_b()));
            }
        },

        updatePicture: function (props) {
            var v = this.view;
            this.lockFor('pictureformat', props.get_Locked());
            this._imgProps = props;
            v.spnPicHeight.setValue(Common.Utils.Metric.fnRecalcFromMM(props.get_Height()).toFixed(2), true);
            v.spnPicWidth.setValue(Common.Utils.Metric.fnRecalcFromMM(props.get_Width()).toFixed(2), true);
            var transp = props.asc_getTransparent(),
                pct = (transp === null || transp === undefined) ? 0 : Math.round(100 - transp / 2.55);
            _.each(v.btnPicTransparency.menu.items, function (it) { it.setChecked(Math.abs(it.value - pct) < 3, true); });
            var ole = props.asc_getPluginGuid() !== null && props.asc_getPluginGuid() !== undefined;
            v.btnPicChange.setDisabled(ole || props.get_Locked());
            v.btnPicCrop.setDisabled(props.get_Locked() || !this.api.asc_canEditCrop());
            var shp = this._picShapeProps, shadow = shp && shp.asc_getShadow && shp.asc_getShadow();
            v.btnPicEffects.menu.items[0].setChecked(!shadow, true);
            v.btnPicEffects.menu.items[1].setChecked(!!shadow, true);
            this.updateArrange(v.picArrange, props);
        },

        updateShape: function (props) {
            var v = this.view;
            this.lockFor('shapeformat', props.get_Locked());
            this._imgProps = props;
            v.spnShpHeight.setValue(Common.Utils.Metric.fnRecalcFromMM(props.get_Height()).toFixed(2), true);
            v.spnShpWidth.setValue(Common.Utils.Metric.fnRecalcFromMM(props.get_Width()).toFixed(2), true);
            var shp = props.get_ShapeProperties();
            if (shp) {
                var vert = props.get_Vert ? props.get_Vert() : null, va = props.get_VerticalTextAlign ? props.get_VerticalTextAlign() : null;
                _.each(v.btnShpDirection.menu.items, function (it) { it.setChecked(it.options.direction === vert, true); });
                _.each(v.btnShpAlignText.menu.items, function (it) { it.setChecked(it.options.valign === va, true); });
                var shadow = shp.asc_getShadow && shp.asc_getShadow();
                v.btnShpEffects.menu.items[0].setChecked(!shadow, true);
                v.btnShpEffects.menu.items[1].setChecked(!!shadow, true);
                v.btnShpEdit.menu.items[1].setDisabled(!this.api.asc_canEditGeometry());
            }
            this.updateArrange(v.shpArrange, props);
        },

        updateArrange: function (arr, props) {
            var wrap = props.get_WrappingStyle(), notflow = !props.get_CanBeFlow(), locked = props.get_Locked();
            arr.position.setDisabled(locked || notflow || props.get_FromGroup());
            _.each(arr.rotate.menu.items, function (it) { it.value !== 'more' && it.setDisabled && it.setDisabled(locked); });
        },

        updateEquation: function () {
            if (!this.api) return;
            var inline = this.api.asc_IsInlineMath && this.api.asc_IsInlineMath();
            this.view.btnEqDisplay.setCaption(inline ? this.view.txtChangeToDisplay : this.view.txtChangeToInline);
            this.view.btnEqDisplay.options.isInline = !!inline;
            this.updateEqInput();
        },

        updateEqInput: function () {
            var eq = this.api.asc_GetMathInputType(), items = this.view.btnEqConvert.menu.items;
            items[5].setChecked(eq === Asc.c_oAscMathInputType.Unicode, true);
            items[6].setChecked(eq === Asc.c_oAscMathInputType.LaTeX, true);
        },

        updateColors: function () {
            var v = this.view;
            if (!v) return;
            var eff = Common.Utils.ThemeColor.getEffectColors(), std = Common.Utils.ThemeColor.getStandartColors();
            if (!eff) return;
            _.each([v.btnShading, v.btnPenColor, v.btnShpFill, v.btnShpOutline, v.btnPicBorder], function (b) {
                if (!b.menu || !_.isObject(b.menu)) { b.setMenu(); b.getPicker(); }
                b.getPicker().updateColors(eff, std);
            });
        },

        // ---- table styles gallery --------------------------------------------------------
        ensureTableStyles: function () {
            if (!this.api || this.view.tableStyles.menuPicker.store.length > 0 || this._state.stylesRequested) return;
            // the previews are shared with the right panel table settings, which must be ready to receive them
            var rm = this.getApplication().getController('RightMenu'),
                ts = rm && rm.rightmenu && rm.rightmenu.tableSettings;
            if (ts && ts._initSettings && ts.createDelayedElements) ts.createDelayedElements();
            if (ts && !ts.mnuTableTemplatePicker) return; // it will generate them itself once the styles are loaded
            this._state.stylesRequested = true;
            this.api.asc_generateTableStylesPreviews();
        },

        onBeginTableStylesPreview: function (count) {
            var v = this.view;
            if (!v) return;
            this._styles = {begin: true, count: count, groups: {
                'menu-table-group-custom': {id: 'menu-table-group-custom', caption: v.txtGroupTable_Custom || 'Custom', index: 0, templateCount: 0},
                'menu-table-group-plain': {id: 'menu-table-group-plain', caption: v.txtGroupTable_Plain || 'Plain Tables', index: 1, templateCount: 0},
                'menu-table-group-grid': {id: 'menu-table-group-grid', caption: v.txtGroupTable_Grid || 'Grid Tables', index: 2, templateCount: 0},
                'menu-table-group-list': {id: 'menu-table-group-list', caption: v.txtGroupTable_List || 'List Tables', index: 3, templateCount: 0},
                'menu-table-group-bordered-and-lined': {id: 'menu-table-group-bordered-and-lined', caption: v.txtGroupTable_BorderedAndLined || 'Bordered & Lined Tables', index: 4, templateCount: 0},
                'menu-table-group-no-name': {id: 'menu-table-group-no-name', caption: '&nbsp', index: 5, templateCount: 0}
            }};
        },

        onAddTableStylesPreview: function (templates) {
            var me = this, v = this.view, st = this._styles;
            if (!v || !st) return;
            var ts = this.getApplication().getController('RightMenu');
            ts = ts && ts.rightmenu && ts.rightmenu.tableSettings; // reuse the localized style names
            var picker = v.tableStyles.menuPicker, field = v.tableStyles.fieldPicker, items = [];
            _.each(templates, function (template) {
                var tip = template.asc_getDisplayName(), group;
                if (template.asc_getType() == 0) {
                    var arr = tip.split(' ');
                    if (/Table Grid/i.test(tip)) group = 'menu-table-group-plain';
                    else if (/Lined|Bordered/i.test(tip)) group = 'menu-table-group-bordered-and-lined';
                    else {
                        group = arr[0] ? 'menu-table-group-' + arr[0].toLowerCase() : '';
                        if (!st.groups[group]) group = 'menu-table-group-no-name';
                    }
                    ts && ['Table Grid', 'Plain Table', 'Grid Table', 'List Table', 'Light', 'Dark', 'Colorful', 'Accent', 'Bordered & Lined', 'Bordered', 'Lined'].forEach(function (item) {
                        var str = 'txtTable_' + item.replace('&', 'And').replace(/ /g, '');
                        ts[str] && (tip = tip.replace(item, ts[str]));
                    });
                } else group = 'menu-table-group-custom';
                if (ts) _.each(st.groups, function (g, k) {
                    var key = 'txtGroupTable_' + {custom: 'Custom', plain: 'Plain', grid: 'Grid', list: 'List', 'bordered-and-lined': 'BorderedAndLined'}[k.replace('menu-table-group-', '')];
                    ts[key] && (g.caption = ts[key]);
                });
                var idx = 0;
                _.each(st.groups, function (g) { if (g.index <= st.groups[group].index) idx += g.templateCount; });
                var obj = {imageUrl: template.asc_getImage(), id: Common.UI.getId(), group: group, templateId: template.asc_getId(), tip: tip, allowSelected: true, selected: false};
                if (st.begin) {
                    st.begin = false;
                    picker.groups.reset(st.groups[group]);
                    picker.groups.comparator = function (item) { return item.get('index'); };
                    picker.store.reset(obj);
                } else {
                    if (st.groups[group].templateCount === 0) picker.groups.add(st.groups[group]);
                    picker.store.add(obj, {at: idx});
                }
                st.groups[group].templateCount++;
            });
        },

        onEndTableStylesPreview: function () {
            this.selectTableStyle(true);
        },

        selectTableStyle: function (force) {
            var v = this.view, combo = v && v.tableStyles;
            if (!combo || !combo.menuPicker.store.length) return;
            var rec = combo.menuPicker.store.findWhere({templateId: this._state.tableStyle});
            combo.suspendEvents();
            if (rec) {
                combo.menuPicker.selectRecord(rec, true);
                combo.fillComboView(rec, true);
            } else if (force) {
                combo.fillComboView(combo.menuPicker.store.at(0), false);
                combo.menuPicker.deselectAll && combo.menuPicker.deselectAll(true);
            }
            combo.resumeEvents();
        },

        // ---- equation structures ---------------------------------------------------------
        ensureEquationStructures: function () {
            var v = this.view, me = this;
            if (v.btnsEqStructures.length) return;
            var store = this.getApplication().getCollection('EquationGroups');
            if (!store.length && this.toolbarCtrl.onMathTypes) this.toolbarCtrl.onMathTypes();
            var btns = v.createEquationStructures(store);
            _.each(btns, function (btn) {
                var onShow = function (menu) {
                    var picker = new Common.UI.DataViewSimple({
                        el: menu.$el.find('.elx-eq-menu'),
                        parentMenu: menu,
                        store: btn.eqGroup.get('groupStore'),
                        scrollAlwaysVisible: true,
                        showLast: false,
                        restoreHeight: 450,
                        itemTemplate: _.template('<div class="item-equation"><div class="equation-icon" style="background-position:<%= posX %>px <%= posY %>px;width:<%= width %>px;height:<%= height %>px;" id="<%= id %>"></div></div>')
                    });
                    picker.on('item:click', function (picker, item, record, e) {
                        record && me.api && me.api.asc_AddMath(record.get('data').equationType);
                        if (e.type !== 'click') menu.hide();
                        Common.NotificationCenter.trigger('edit:complete', me.toolbar);
                    });
                    menu.off('show:after', onShow);
                };
                btn.menu.on('show:after', onShow);
            });
            this.toolbar.clearActiveData && this.toolbar.clearActiveData('equation');
            this.toolbar.processPanelVisible(null, true);
        },

        // ---- events --------------------------------------------------------------------------
        bindEvents: function () {
            var me = this, v = this.view, api = function () { return me.api; };
            var done = function () { Common.NotificationCenter.trigger('edit:complete', me.toolbar); };
            var tbl = function (fill) {
                if (!me.api) return;
                var p = new Asc.CTableProp();
                fill(p);
                me.api.tblApply(p);
                done();
            };
            var img = function (fill) {
                if (!me.api) return;
                var p = new Asc.asc_CImgProperty();
                fill(p);
                me.api.ImgApply(p);
                done();
            };
            var shape = function (fill) {
                img(function (p) { var s = new Asc.asc_CShapeProperty(); fill(s); p.put_ShapeProperties(s); });
            };

            // Table Design
            var look = function (setter, ch) {
                ch.on('change', function (field) {
                    tbl(function (p) {
                        var l = (me._tableProps && me._tableProps.get_TableLook()) || new Asc.CTablePropLook();
                        l[setter](field.getValue() == 'checked');
                        p.put_TableLook(l);
                    });
                });
            };
            look('put_FirstRow', v.chHeader); look('put_LastRow', v.chTotal); look('put_BandHor', v.chBanded);
            look('put_FirstCol', v.chFirst); look('put_LastCol', v.chLast); look('put_BandVer', v.chColBanded);
            v.tableStyles.on('click', function (combo, record) {
                record && tbl(function (p) { p.put_TableStyle(record.get('templateId')); });
            });
            v.tableStyles.openButton && v.tableStyles.openButton.menu.on('show:before', function () { me.ensureTableStyles(); });
            v.btnShading.on('color:select', function (btn, color) {
                tbl(function (p) {
                    var bg = new Asc.CBackground();
                    if (color == 'transparent') bg.put_Value(1);
                    else { bg.put_Value(0); bg.put_Color(Common.Utils.ThemeColor.getRgbColor(color)); }
                    p.put_CellsBackground(bg);
                    p.put_CellSelect(true);
                });
            });
            v.btnShading.on('click', function (btn) {
                v.btnShading.currentColor !== undefined && v.btnShading.trigger('color:select', v.btnShading, v.btnShading.color || 'transparent');
            });
            v.cmbBorderSize.on('selected', function (combo, record) { v.BorderSize = record.value; });
            v.btnBorders.menu.on('item:click', function (menu, item) {
                if (item.value === 'advanced') { me.openTableAdvanced(2); return; }
                v.btnBorders.options.lastBorder = item.value;
                me.applyCellBorders(item.value);
            });
            v.btnBorders.on('click', function () { me.applyCellBorders(v.btnBorders.options.lastBorder); });

            // Table Layout
            v.btnTblSelect.menu.on('item:click', function (menu, item) {
                var a = me.api; if (!a) return;
                ({cell: a.selectCell, column: a.selectColumn, row: a.selectRow, table: a.selectTable})[item.value].call(a);
                done();
            });
            v.btnTblGridlines.on('click', function (btn) {
                var state = btn.pressed;
                Common.localStorage.setItem('de-show-tableline', state);
                me.api && me.api.put_ShowTableEmptyLine(state);
                var mnu = me.toolbar.mnuNonPrinting;
                mnu && mnu.items[1] && mnu.items[1].setChecked(state, true);
                done();
            });
            v.btnTblProps.on('click', function () { me.openTableAdvanced(0); });
            v.btnTblDraw.on('click', function (btn) { me.api && me.api.SetTableDrawMode(btn.pressed); });
            v.btnTblEraser.on('click', function (btn) { me.api && me.api.SetTableEraseMode(btn.pressed); });
            v.btnTblDelete.menu.on('item:click', function (menu, item) {
                var a = me.api; if (!a) return;
                if (item.value === 'cells') me.dh().onCellsRemove();
                else ({column: a.remColumn, row: a.remRow, table: a.remTable})[item.value].call(a);
                done();
            });
            v.btnTblInsAbove.on('click', function () { me.api && me.api.addRowAbove(); done(); });
            v.btnTblInsBelow.on('click', function () { me.api && me.api.addRowBelow(); done(); });
            v.btnTblInsLeft.on('click', function () { me.api && me.api.addColumnLeft(); done(); });
            v.btnTblInsRight.on('click', function () { me.api && me.api.addColumnRight(); done(); });
            v.btnTblMerge.on('click', function () { me.api && me.api.MergeCells(); done(); });
            v.btnTblSplit.on('click', function () { me.dh().onTableSplit(); });
            v.btnTblAutofit.menu.on('item:click', function (menu, item) {
                tbl(function (p) {
                    if (item.value === 'fixed') p.put_TableLayout(Asc.c_oAscTableLayout.Fixed);
                    else {
                        p.put_TableLayout(Asc.c_oAscTableLayout.AutoFit);
                        item.value === 'window' && p.put_Width(-100);
                    }
                });
            });
            v.spnTblHeight.on('change', function (field) { tbl(function (p) { p.put_RowHeight(Common.Utils.Metric.fnRecalcToMM(field.getNumberValue())); }); });
            v.spnTblWidth.on('change', function (field) { tbl(function (p) { p.put_ColumnWidth(Common.Utils.Metric.fnRecalcToMM(field.getNumberValue())); }); });
            v.btnTblDistRows.on('click', function () { me.api && me.api.asc_DistributeTableCells(false); done(); });
            v.btnTblDistCols.on('click', function () { me.api && me.api.asc_DistributeTableCells(true); done(); });
            _.each(v.btnsCellAlign, function (b) {
                b.on('click', function () {
                    if (!me.api) return;
                    me.api.put_PrAlign(b.options.halign);
                    var p = new Asc.CTableProp();
                    p.put_CellsVAlign(b.options.valign);
                    me.api.tblApply(p);
                    done();
                });
            });
            v.btnTblDirection.menu.on('item:click', function (menu, item) { tbl(function (p) { p.put_CellsTextDirection(item.options.direction); }); });
            v.btnTblMargins.on('click', function () { me.openTableAdvanced(1); });
            v.btnTblRepeatHeader.on('click', function (btn) { tbl(function (p) { p.put_RowsInHeader(btn.pressed); }); });
            v.btnTblToText.on('click', function () {
                (new DE.Views.TableToTextDialog({
                    handler: function (dlg, result) {
                        if (result == 'ok') {
                            var s = dlg.getSettings();
                            me.api.asc_ConvertTableToText(s.type, s.separator, s.nested);
                        }
                        done();
                    }
                })).show();
            });
            v.btnTblFormula.on('click', function () {
                (new DE.Views.TableFormulaDialog({
                    api: me.api,
                    bookmarks: me.api.asc_GetBookmarksManager(),
                    handler: function (result, value) {
                        result == 'ok' && me.api.asc_AddTableFormula(value);
                        done();
                    }
                })).show();
            });

            // Picture Format
            v.btnPicTransparency.menu.on('item:click', function (menu, item) {
                img(function (p) { p.asc_putTransparent((100 - item.value) * 2.55); });
            });
            v.btnPicChange.menu.on('item:click', function (menu, item) { me.changePicture(item.value); });
            var resetPic = function (withSize) {
                img(function (p) {
                    p.put_ResetCrop(true);
                    if (withSize) {
                        var sz = me.api.asc_getCropOriginalImageSize();
                        sz && p.put_Width(sz.get_ImageWidth());
                        sz && p.put_Height(sz.get_ImageHeight());
                        p.put_Rot(0);
                    }
                });
            };
            v.btnPicReset.on('click', function () { resetPic(false); });
            v.btnPicReset.menu.on('item:click', function (menu, item) { resetPic(item.value === 'size'); });
            v.btnPicAltText.on('click', function () { me.openImageAdvanced(12); });
            v.btnPicCrop.on('click', function (btn) {
                if (!me.api) return;
                btn.pressed ? me.api.asc_startEditCrop() : me.api.asc_endEditCrop();
                done();
            });
            v.btnPicCrop.menu.on('item:click', function (menu, item) {
                if (!me.api) return;
                if (item.value === 1) me.api.asc_cropFill();
                else if (item.value === 2) me.api.asc_cropFit();
                else item.checked ? me.api.asc_startEditCrop() : me.api.asc_endEditCrop();
                done();
            });
            v.btnPicCrop.menu.on('show:before', function (menu) { menu.items[0].setChecked(v.btnPicCrop.pressed, true); });
            this.bindSize(v.spnPicWidth, v.spnPicHeight, true);
            this.bindArrange(v.picArrange);

            // Shape Format
            this.bindShapePicker();
            v.btnShpEdit.menu.on('item:click', function (menu, item) {
                item.value === 'points' && me.api && me.api.asc_editPointsGeometry();
            });
            v.btnShpEdit.menu.items[0].menu.on('show:before', function (menu) { me.fillChangeShape(menu); });
            v.btnShpFill.on('color:select', function (btn, color) {
                shape(function (s) {
                    var fill = new Asc.asc_CShapeFill();
                    if (color == 'transparent') { fill.put_type(Asc.c_oAscFill.FILL_TYPE_NOFILL); fill.put_fill(null); }
                    else {
                        fill.put_type(Asc.c_oAscFill.FILL_TYPE_SOLID);
                        fill.put_fill(new Asc.asc_CFillSolid());
                        fill.get_fill().put_color(Common.Utils.ThemeColor.getRgbColor(color));
                    }
                    s.put_fill(fill);
                });
            });
            v.btnShpFill.on('click', function () { v.btnShpFill.color !== undefined && v.btnShpFill.trigger('color:select', v.btnShpFill, v.btnShpFill.color); });
            var stroke = function (color, width) {
                shape(function (s) {
                    var st = new Asc.asc_CStroke();
                    if (color == 'transparent') st.put_type(Asc.c_oAscStrokeType.STROKE_NONE);
                    else {
                        st.put_type(Asc.c_oAscStrokeType.STROKE_COLOR);
                        st.put_color(Common.Utils.ThemeColor.getRgbColor(color || '000000'));
                        st.put_width((width || me._state.strokeWidth || 1) * 25.4 / 72);
                    }
                    s.put_stroke(st);
                });
            };
            _.each([v.btnShpOutline, v.btnPicBorder], function (b) {
                b.on('color:select', function (btn, color) { stroke(color); });
                b.on('click', function () { stroke(b.color || '000000'); });
                b.mnuWeight.menu.on('item:click', function (menu, item) {
                    me._state.strokeWidth = item.value;
                    stroke(b.color && b.color !== 'transparent' ? b.color : '000000', item.value);
                });
            });
            var effects = function (menu, item) {
                if (item.value === 'options') {
                    var props = (me._imgProps && me._imgProps.get_ShapeProperties()) || me._picShapeProps;
                    (new Common.Views.ShapeShadowDialog({
                        api: me.api,
                        shadowProps: props ? props.asc_getShadow() : null,
                        methodApplySettings: function (shapeProps) {
                            var p = new Asc.asc_CImgProperty();
                            p.put_ShapeProperties(shapeProps);
                            me.api.ImgApply(p);
                        },
                        handler: function () { done(); }
                    })).show();
                    return;
                }
                shape(function (s) {
                    if (item.value === 'none') s.asc_putShadow(null);
                    else { var sh = new Asc.asc_CShadowProperty(); sh.putPreset('t'); s.asc_putShadow(sh); }
                });
            };
            v.btnShpEffects.menu.on('item:click', effects);
            v.btnPicEffects.menu.on('item:click', effects);
            v.btnShpDirection.menu.on('item:click', function (menu, item) { img(function (p) { p.put_Vert(item.options.direction); }); });
            v.btnShpAlignText.menu.on('item:click', function (menu, item) { img(function (p) { p.put_VerticalTextAlign(item.options.valign); }); });
            v.btnShpAltText.on('click', function () { me.openImageAdvanced(12); });
            this.bindSize(v.spnShpWidth, v.spnShpHeight, false);
            this.bindArrange(v.shpArrange);

            // Equation
            v.btnEqNew.on('click', function () { me.api && me.api.asc_AddMath(); done(); });
            v.btnEqConvert.menu.on('item:click', function (menu, item) {
                if (!me.api) return;
                if (item.options.type === 'input') {
                    me.api.asc_SetMathInputType(item.value);
                    Common.localStorage.setBool('de-equation-input-latex', item.value === Asc.c_oAscMathInputType.LaTeX);
                } else if (item.options.type === 'view') me.api.asc_ConvertMathView(item.value.linear, item.value.all);
                done();
            });
            v.btnEqDisplay.on('click', function () {
                me.api && me.api.asc_ConvertMathDisplayMode(!v.btnEqDisplay.options.isInline);
                done();
            });
        },

        bindShapePicker: function () {
            var me = this, v = this.view, btn = v.btnShpShapes;
            // the Insert tab Shapes gallery holds a picker that cannot be cloned: build our own
            btn.menu.on('show:before', function (menu) {
                if (menu._elxPicker) return;
                menu.removeAll();
                var item = new Common.UI.MenuItem({template: _.template('<div class="menu-insertshape elx-shape-picker"></div>')});
                menu.addItem(item, true);
                var recents = Common.localStorage.getItem('de-recent-shapes');
                var picker = new Common.UI.DataViewShape({
                    el: menu.$el.find('.elx-shape-picker'),
                    itemTemplate: _.template('<div class="item-shape" id="<%= id %>"><svg width="20" height="20" class=\"icon uni-scale\"><use xlink:href=\"#svg-icon-<%= data.shapeType %>\"></use></svg></div>'),
                    groups: me.getApplication().getCollection('ShapeGroups'),
                    parentMenu: menu,
                    restoreHeight: 652,
                    textRecentlyUsed: me.toolbarCtrl.textRecentlyUsed,
                    recentShapes: recents ? JSON.parse(recents) : null
                });
                menu.setInnerMenu([{menu: picker, index: 0}]);
                picker.on('item:click', function (picker, item, record, e) {
                    record && me.toolbarCtrl._addAutoshape && me.toolbarCtrl._addAutoshape(true, record.get('data').shapeType);
                    if (e.type !== 'click') menu.hide();
                    Common.NotificationCenter.trigger('edit:complete', me.toolbar);
                });
                menu._elxPicker = picker;
            });
        },

        fillChangeShape: function (menu) {
            var me = this;
            if (menu._elxPicker) return;
            var item = new Common.UI.MenuItem({template: _.template('<div class="menu-insertshape elx-change-shape-picker"></div>')});
            menu.addItem(item, true);
            var picker = new Common.UI.DataViewShape({
                el: menu.$el.find('.elx-change-shape-picker'),
                itemTemplate: _.template('<div class="item-shape" id="<%= id %>"><svg width="20" height="20" class=\"icon uni-scale\"><use xlink:href=\"#svg-icon-<%= data.shapeType %>\"></use></svg></div>'),
                groups: me.getApplication().getCollection('ShapeGroups'),
                parentMenu: menu,
                restoreHeight: 652,
                hideTextRect: true,
                hideLines: true
            });
            menu.setInnerMenu([{menu: picker, index: 0}]);
            picker.on('item:click', function (picker, item, record, e) {
                record && me.api && me.api.ChangeShapeType(record.get('data').shapeType);
                if (e.type !== 'click') me.view.btnShpEdit.menu.hide();
                Common.NotificationCenter.trigger('edit:complete', me.toolbar);
            });
            menu._elxPicker = picker;
        },

        bindSize: function (spnW, spnH, keepRatio) {
            var me = this;
            var apply = function (isWidth, field) {
                var props = me._imgProps;
                if (!me.api || !props) return;
                var w = props.get_Width(), h = props.get_Height(), val = Common.Utils.Metric.fnRecalcToMM(field.getNumberValue()),
                    p = new Asc.asc_CImgProperty(),
                    lock = keepRatio || (props.asc_getLockAspect && props.asc_getLockAspect());
                if (isWidth) { p.put_Width(val); lock && w > 0 && p.put_Height(h * val / w); }
                else { p.put_Height(val); lock && h > 0 && p.put_Width(w * val / h); }
                me.api.ImgApply(p);
            };
            spnW.on('change', function (field) { apply(true, field); });
            spnH.on('change', function (field) { apply(false, field); });
        },

        bindArrange: function (arr) {
            var me = this;
            arr.position.menu.on('item:click', function (menu, item) { me.onPositionClick(item); });
            arr.rotate.menu.on('item:click', function (menu, item) { me.onRotateClick(item); });
        },

        // also used by the Layout tab (DE.Controllers.RibbonTabs)
        onPositionClick: function (item) {
            if (item.value === 'more') { this.openImageAdvanced(9); return; }
            this.applyPosition(item.value);
        },

        onRotateClick: function (item) {
            if (!this.api) return;
            if (item.value === 'more') { this.openImageAdvanced(7); return; }
            var p = new Asc.asc_CImgProperty();
            if (item.value === 'r90' || item.value === 'r270') p.asc_putRotAdd((item.value === 'r90' ? 90 : 270) * Math.PI / 180);
            else if (item.value === 'fh') p.asc_putFlipHInvert(true);
            else p.asc_putFlipVInvert(true);
            this.api.ImgApply(p);
            Common.NotificationCenter.trigger('edit:complete', this.toolbar);
        },

        // Word's Position gallery: square wrapping, aligned on the margins
        applyPosition: function (index) {
            if (!this.api) return;
            var p = new Asc.asc_CImgProperty();
            if (index === 'inline') {
                p.put_WrappingStyle(Asc.c_oAscWrapStyle2.Inline);
            } else {
                var hAlign = [Asc.c_oAscAlignH.Left, Asc.c_oAscAlignH.Center, Asc.c_oAscAlignH.Right][index % 3],
                    vAlign = [Asc.c_oAscAlignV.Top, Asc.c_oAscAlignV.Center, Asc.c_oAscAlignV.Bottom][Math.floor(index / 3)];
                var props = this._imgProps;
                if (!props || props.get_WrappingStyle() === Asc.c_oAscWrapStyle2.Inline)
                    p.put_WrappingStyle(Asc.c_oAscWrapStyle2.Square);
                var ph = new Asc.CImagePositionH();
                ph.put_UseAlign(true); ph.put_Percent(false); ph.put_RelativeFrom(Asc.c_oAscRelativeFromH.Margin); ph.put_Align(hAlign);
                var pv = new Asc.CImagePositionV();
                pv.put_UseAlign(true); pv.put_Percent(false); pv.put_RelativeFrom(Asc.c_oAscRelativeFromV.Margin); pv.put_Align(vAlign);
                p.put_PositionH(ph);
                p.put_PositionV(pv);
            }
            this.api.ImgApply(p);
            Common.NotificationCenter.trigger('edit:complete', this.toolbar);
        },

        applyCellBorders: function (border) {
            if (!this.api) return;
            var v = this.view, borders = new Asc.CBorders(), visible = border !== '';
            var set = function (getter, putter, key) {
                if (border.indexOf(key) > -1 || !visible) {
                    var b = new Asc.asc_CTextBorder();
                    if (visible && v.BorderSize > 0) {
                        b.put_Value(1);
                        b.put_Size(parseFloat(v.BorderSize) * 25.4 / 72.0);
                        var color;
                        if (v.btnPenColor.isAutoColor()) { color = new Asc.asc_CColor(); color.put_auto(true); }
                        else color = Common.Utils.ThemeColor.getRgbColor(v.btnPenColor.color);
                        b.put_Color(color);
                    } else b.put_Value(0);
                    borders[putter](b);
                }
            };
            set('get_Left', 'put_Left', 'l');
            set('get_Top', 'put_Top', 't');
            set('get_Right', 'put_Right', 'r');
            set('get_Bottom', 'put_Bottom', 'b');
            set('get_InsideV', 'put_InsideV', 'c');
            set('get_InsideH', 'put_InsideH', 'm');
            var p = new Asc.CTableProp();
            p.put_CellBorders(borders);
            p.put_CellSelect(true);
            this.api.tblApply(p);
            Common.NotificationCenter.trigger('edit:complete', this.toolbar);
        },

        changePicture: function (value) {
            var me = this;
            if (value === 1) {
                (new Common.Views.ImageFromUrlDialog({
                    handler: function (result, url) {
                        if (result == 'ok' && me.api) {
                            var u = url.replace(/ /g, '');
                            if (!_.isEmpty(u)) { var p = new Asc.asc_CImgProperty(); p.put_ImageUrl(u); me.api.ImgApply(p); }
                        }
                        Common.NotificationCenter.trigger('edit:complete', me.toolbar);
                    }
                })).show();
            } else if (value === 2) {
                Common.NotificationCenter.trigger('storage:image-load', 'change');
            } else {
                me.api && me.api.ChangeImageFromFile();
                Common.NotificationCenter.trigger('edit:complete', me.toolbar);
            }
        },

        dh: function () {
            return this.getApplication().getController('DocumentHolder');
        },

        // Picture and shape "Layout" dialog, opened on a given page (see ImageSettingsAdvanced items)
        openImageAdvanced: function (category) {
            var me = this, api = this.api;
            if (!api) return;
            var sel = api.getSelectedElements();
            for (var i = sel.length - 1; i >= 0; i--) {
                if (sel[i].get_ObjectType() !== Asc.c_oAscTypeSelectElement.Image) continue;
                var value = sel[i].get_ObjectValue(), sizeOriginal;
                if (!value.get_ChartProperties() && !value.get_ShapeProperties() && value.get_ImageUrl()) {
                    var s = api.asc_getCropOriginalImageSize();
                    s && (sizeOriginal = {width: s.get_ImageWidth(), height: s.get_ImageHeight()});
                }
                var win = new DE.Views.ImageSettingsAdvanced({
                    imageProps: value,
                    sizeOriginal: sizeOriginal,
                    api: api,
                    sectionProps: api.asc_GetSectionProps(),
                    chartSettings: null,
                    handler: function (result, v) {
                        result == 'ok' && api.ImgApply(v.imageProps);
                        Common.NotificationCenter.trigger('edit:complete', me.toolbar);
                    }
                });
                win.show();
                category !== undefined && win.setActiveCategory(category);
                return;
            }
        },

        // Table Properties dialog, opened on a given page: 0 table, 1 cell, 2 borders
        openTableAdvanced: function (category) {
            var me = this, api = this.api;
            if (!api) return;
            var sel = api.getSelectedElements();
            for (var i = sel.length - 1; i >= 0; i--) {
                if (sel[i].get_ObjectType() !== Asc.c_oAscTypeSelectElement.Table) continue;
                var value = sel[i].get_ObjectValue();
                var win = new DE.Views.TableSettingsAdvanced({
                    tableStylerRows: (value.get_CellBorders().get_InsideH() === null && value.get_CellSelect() == true) ? 1 : 2,
                    tableStylerColumns: (value.get_CellBorders().get_InsideV() === null && value.get_CellSelect() == true) ? 1 : 2,
                    tableProps: value,
                    borderProps: me._borderAdvancedProps,
                    sectionProps: api.asc_GetSectionProps(),
                    handler: function (result, v) {
                        if (result == 'ok') { me._borderAdvancedProps = v.borderProps; api.tblApply(v.tableProps); }
                        Common.NotificationCenter.trigger('edit:complete', me.toolbar);
                    }
                });
                win.show();
                category !== undefined && win.setActiveCategory(category);
                return;
            }
        },

        onGroupLauncher: function (group) {
            _.isArray(group) && (group = group[0]);
            switch (group) {
            case 'cells': this.dh().onCellsAdd(); break;
            case 'cellsize': this.openTableAdvanced(1); break;
            case 'cellmargins': this.openTableAdvanced(0); break;
            case 'borders': this.openTableAdvanced(2); break;
            case 'picsize': this.openImageAdvanced(5); break;
            case 'shpsize': this.openImageAdvanced(6); break;
            case 'picturestyles':
            case 'shapestyles':
                var rm = this.getApplication().getController('RightMenu');
                rm && rm.onRightMenuOpen(Common.Utils.documentSettingsType.Shape);
                break;
            }
        }
    }, DE.Controllers.ContextTabs || {}));
});

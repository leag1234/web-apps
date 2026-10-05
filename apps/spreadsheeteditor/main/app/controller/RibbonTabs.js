/*
 * SPDX-FileCopyrightText: 2026 ElixirOffice contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 *
 * SSE.Controllers.RibbonTabs: an Excel-like ribbon for the spreadsheet editor.
 *
 * - Toolbar 'render:before': the group label logic (Common.UI.RibbonGroupLabels)
 *   is mixed into the toolbar, the stock buttons Excel shows large or with a
 *   caption are restyled, and the controls Excel has and the stock ribbon
 *   lacks are created (SSE.Views.RibbonTabs).
 * - Toolbar 'render:after': those controls are rendered into their slots
 *   (Toolbar.template), the ribbon can be collapsed / pinned and its groups
 *   collapse like Office (Common.UI.RibbonCollapse).
 * - app ready: tab names and KeyTips of Excel, Protection merged into Review.
 * Group launchers ('group:launcher') open the matching dialog or pane.
 */

define([
    'core',
    'spreadsheeteditor/main/app/view/RibbonTabs',
    'common/main/lib/component/RibbonGroupLabels',
    'common/main/lib/component/RibbonCollapse'
], function () {
    'use strict';

    SSE.Controllers.RibbonTabs = Backbone.Controller.extend(_.extend({
        models: [],
        collections: [],
        views: ['RibbonTabs'],

        initialize: function () {
            var me = this;
            this.addListeners({
                'Toolbar': {
                    'render:before': function (toolbar) { me.onToolbarBeforeRender(toolbar); },
                    'render:after': function (toolbar) { me.onToolbarAfterRender(toolbar); },
                    'group:launcher': function (key) { me.onLauncher(key); }
                }
            });
            Common.NotificationCenter.on('app:ready', this.onAppReady.bind(this));
        },

        appOptions: function () {
            var main = this.getApplication().getController('Main');
            return (main && main.appOptions) || {};
        },

        ctrl: function (name) { return this.getApplication().getController(name); },

        onToolbarBeforeRender: function (toolbar) {
            var mode = this.appOptions();
            if (!mode.isEdit || mode.isEditDiagram || mode.isEditMailMerge || mode.isEditOle || this.view) return;
            this.toolbar = toolbar;
            var v = this.view = this.createView('RibbonTabs', {toolbar: toolbar, mode: mode});

            // group labels and launchers, like DE.Views.Toolbar
            _.extend(toolbar, Common.UI.RibbonGroupLabels);
            toolbar.groupTags = v.groupTags;
            _.each(_.keys(v.constructor.prototype), function (k) {
                if (/^txtGroup/.test(k) || k === 'tipGroupLauncher') toolbar[k] = v[k];
            });

            this.restyle(toolbar);
            v.createControls();
        },

        // stock buttons in Excel's size and words
        restyle: function (tb) {
            var v = this.view;
            var set = function (b, o) {
                if (!b) return;
                if (o.cls !== undefined) b.cls = b.options.cls = o.cls;
                if (o.iconCls !== undefined) b.iconCls = b.options.iconCls = o.iconCls;
                if (o.caption !== undefined) b.caption = b.options.caption = o.caption;
                if (o.cls && /x-huge/.test(o.cls)) {
                    b.options.dataHintDirection = 'bottom';
                    b.options.dataHintOffset = 'small';
                }
            };
            var big = 'btn-toolbar x-huge icon-top';
            set(tb.btnPaste, {cls: big, iconCls: 'toolbar__icon btn-big-paste', caption: v.txtPaste});
            set(tb.btnCondFormat, {cls: big, iconCls: 'toolbar__icon btn-big-elx-cond-format', caption: v.txtCondFormat});
            set(tb.btnTableTemplate, {cls: big, iconCls: 'toolbar__icon btn-big-elx-format-table', caption: v.txtFormatAsTable});
            set(tb.btnAddCell, {cls: big, iconCls: 'toolbar__icon btn-big-elx-insert-cells', caption: v.txtInsert});
            set(tb.btnDeleteCell, {cls: big, iconCls: 'toolbar__icon btn-big-elx-delete-cells', caption: v.txtDelete});
            set(tb.btnFormatCell, {caption: v.txtFormat});
            set(tb.btnWrap, {cls: 'btn-toolbar elx-captioned', caption: v.txtWrapText});
            set(tb.btnMerge, {cls: 'btn-toolbar elx-captioned', caption: v.txtMergeCenter});
            set(tb.btnInsertFormula, {cls: 'btn-toolbar elx-captioned', caption: v.txtAutoSum});
            set(tb.btnFillNumbers, {cls: 'btn-toolbar elx-captioned', caption: v.txtFill});
            set(tb.btnClearStyle, {cls: 'btn-toolbar elx-captioned', caption: v.txtClear});
        },

        onToolbarAfterRender: function (toolbar) {
            if (!this.view) return;
            var v = this.view, me = this;
            v.renderControls(toolbar.$el);
            Array.prototype.push.apply(toolbar.lockControls, v.lockedControls);
            this.bindControls();
            this.setupRibbonBehaviour(toolbar);
            this.api = this.ctrl('Viewport').getApi ? this.ctrl('Viewport').getApi() : null;
            this.api && this.api.asc_registerCallback('asc_onSelectionChanged', _.bind(this.onSelectionChanged, this));
            // the empty Insert > Icons slot (no icon gallery in the spreadsheet yet)
            toolbar.$el.find('#slot-elx-ins-icons:empty').remove();
            setTimeout(function () { me.toolbar.processPanelVisible && me.toolbar.processPanelVisible(null, true); }, 0);
        },

        onAppReady: function () {
            if (!this.view) return;
            var me = this;
            Common.Utils.lockControls(Common.enumLock.disableOnStart, false, {array: this.view.lockedControls});
            this.bindRibbonShortcuts();
            setTimeout(function () {
                me.arrangeTabs();
                me.mergeProtection();
                me.reorderRuns('review', ['proofing', 'comments', 'protect', 'collaboration', 'chat', 'history']);
                me.reorderRuns('view', ['sheetview', 'workbookviews', 'show', 'zoom', 'window', 'interface', 'appearance', 'macros']);
                me.captionDataTab();
                me.toolbar.processPanelVisible && me.toolbar.processPanelVisible(null, true);
            }, 0);
        },

        // ---- tabs: Excel's names and KeyTips -----------------------------------------
        arrangeTabs: function () {
            var tb = this.toolbar, v = this.view;
            if (!tb || !tb.$el) return;
            var names = {layout: v.txtTabPageLayout, formula: v.txtTabFormulas, review: v.txtTabReview, pivot: v.txtTabPivot},
                keys = {file: 'F', home: 'H', ins: 'N', draw: 'JI', layout: 'P', formula: 'M', data: 'A', review: 'R', view: 'W', pivot: 'JT', tabledesign: 'JT'};
            tb.$el.find('.ribtab > a[data-tab]').each(function () {
                var $a = $(this), tab = $a.attr('data-tab');
                names[tab] && $a.text(names[tab]);
                keys[tab] && $a.attr('data-hint-title', keys[tab]);
                // contextual tabs: at the right, in the accent colour, like Excel
                /^(pivot|tabledesign|chartdesign|chartformat|shapeformat|pictureformat)$/.test(tab) && $a.closest('.ribtab').addClass('contextual');
            });
        },

        // Excel's Review tab holds Protect Sheet / Protect Workbook / Allow Edit Ranges
        mergeProtection: function () {
            var tb = this.toolbar;
            if (!tb || !tb.$el) return;
            var $review = tb.$el.find('.box-panels > .panel[data-tab=review]'),
                $protect = tb.$el.find('.box-panels > .panel[data-tab=protect]');
            if (!$review.length || !$protect.length || $protect.hasClass('elx-merged')) return;
            if (tb.isTabVisible && !tb.isTabVisible('review')) return;
            $protect.addClass('elx-merged');
            $review.append('<div class="separator long"></div>');
            $review.append($protect.children());
            tb.setVisible('protect', false);
            tb.tagPanelGroups($review);
        },

        // Put the groups of a panel built by another controller in Excel's order. A
        // "run" is the groups between two separators; its key is the data-group of
        // its first tagged group (Common.UI.RibbonGroupLabels.tagPanelGroups).
        reorderRuns: function (tab, order) {
            var tb = this.toolbar;
            var $panel = tb && tb.$el && tb.$el.find('.box-panels > .panel[data-tab=' + tab + ']');
            if (!$panel || !$panel.length) return;
            tb.tagPanelGroups && tb.tagPanelGroups($panel);
            var runs = [], cur = [];
            $panel.children().each(function () {
                var $c = $(this);
                if ($c.hasClass('separator') && $c.hasClass('long')) {
                    runs.push({items: cur, sep: $c});
                    cur = [];
                } else cur.push($c);
            });
            cur.length && runs.push({items: cur, sep: null});
            var keyOf = function (run) {
                for (var i = 0; i < run.items.length; i++) {
                    var k = run.items[i].attr('data-group');
                    if (k) return k;
                }
                return '';
            };
            var rank = function (run) { var i = order.indexOf(keyOf(run)); return i < 0 ? order.length : i; };
            var sorted = _.sortBy(runs.map(function (r, i) { return {r: r, i: i}; }), function (o) { return rank(o.r) * 1000 + o.i; });
            _.each(sorted, function (o, i) {
                _.each(o.r.items, function ($c) { $panel.append($c); });
                var $sep = o.r.sep || $('<div class="separator long"></div>');
                $panel.append($sep);
                i === sorted.length - 1 && $sep.css('display', 'none');
            });
        },

        // Data > Sort & Filter: captions like Excel
        captionDataTab: function () {
            var tb = this.toolbar, v = this.view;
            var inData = function (arr) {
                return _.find(arr || [], function (b) { return b && b.cmpEl && b.cmpEl.closest('.panel[data-tab=data]').length; });
            };
            var f = inData(tb.btnsSetAutofilter), c = inData(tb.btnsClearAutofilter);
            f && f.setCaption && f.setCaption(v.txtFilter);
            c && c.setCaption && c.setCaption(v.txtClearFilter);
        },

        // ---- new controls ---------------------------------------------------------------
        homeBtn: function (arr) {
            return _.find(arr || [], function (b) { return b && b.cmpEl && b.cmpEl.closest('.panel[data-tab=home]').length; }) || (arr && arr[0]);
        },

        bindControls: function () {
            var me = this, v = this.view;
            v.btnIndentDec.on('click', function () { me.changeIndent(-1); });
            v.btnIndentInc.on('click', function () { me.changeIndent(1); });

            var sources = function (value) {
                var tb = me.toolbar;
                switch (value) {
                    case 'sort-down': return me.homeBtn(tb.btnsSortDown);
                    case 'sort-up': return me.homeBtn(tb.btnsSortUp);
                    case 'sort-custom': return _.isArray(tb.btnCustomSort) ? tb.btnCustomSort[0] : tb.btnCustomSort;
                    case 'set-filter': return me.homeBtn(tb.btnsSetAutofilter);
                    case 'clear-filter': return me.homeBtn(tb.btnsClearAutofilter);
                }
            };
            v.btnSortFilter.menu.on('show:before', function (menu) {
                _.each(menu.items, function (it) {
                    if (!it.value) return;
                    var src = sources(it.value);
                    it.setDisabled(!src || src.isDisabled());
                    it.value === 'set-filter' && it.setChecked(!!(src && src.pressed), true);
                });
            });
            v.btnSortFilter.menu.on('item:click', function (menu, item) {
                var src = sources(item.value);
                if (!src || src.isDisabled()) return;
                if (item.value === 'set-filter' && src.options.enableToggle) src.toggle(!src.pressed, true);
                src.trigger('click', src);
            });

            v.btnFindSelect.menu.on('item:click', function (menu, item) { me.onFindSelect(item.value); });
        },

        changeIndent: function (delta) {
            var api = this.api || (this.toolbar && this.toolbar.api);
            if (!api) return;
            var info = api.asc_getCellInfo(), xfs = info && info.asc_getXfs && info.asc_getXfs(),
                cur = xfs && xfs.asc_getIndent ? (xfs.asc_getIndent() || 0) : 0,
                val = Math.max(0, Math.min(250, cur + delta));
            if (val !== cur) api.asc_setCellIndent(val);
            Common.NotificationCenter.trigger('edit:complete', this.toolbar);
        },

        onSelectionChanged: function (info) {
            var v = this.view;
            if (!v || !info) return;
            var xfs = info.asc_getXfs && info.asc_getXfs(),
                indent = xfs && xfs.asc_getIndent ? xfs.asc_getIndent() : 0;
            v.btnIndentDec.setDisabled(v.btnIndentDec.isDisabled() || !indent);
        },

        onFindSelect: function (value) {
            var lmc = this.ctrl('LeftMenu'), me = this;
            switch (value) {
                case 'find':
                case 'replace':
                    if (!lmc) return;
                    lmc.onShowHideSearch(true);
                    if (value === 'replace') {
                        var panel = lmc.leftMenu && lmc.leftMenu.panelSearch;
                        setTimeout(function () { panel && panel.focus && panel.focus('replace'); }, 100);
                    }
                    break;
                case 'goto':
                    // Excel's Go To: the reference box (Name Box) takes the address
                    setTimeout(function () {
                        var $nb = $('#ce-cell-name');
                        me.api && me.api.asc_enableKeyEvents(false);
                        $nb.focus().select();
                    }, 50);
                    break;
                case 'selectall':
                    var b = this.toolbar && this.toolbar.btnSelectAll;
                    b && !b.isDisabled() && b.trigger('click', b);
                    break;
            }
        },

        // ---- group launchers --------------------------------------------------------------
        onLauncher: function (key) {
            var tbc = this.ctrl('Toolbar'), rm = this.ctrl('RightMenu');
            switch (key) {
                case 'font':
                case 'alignment':
                    rm && rm.onRightMenuOpen && rm.onRightMenuOpen(Common.Utils.documentSettingsType.Cell);
                    break;
                case 'number':
                    tbc && tbc.onCustomNumberFormat && tbc.onCustomNumberFormat();
                    break;
                case 'charts':
                    tbc && tbc.onChartRecommendedClick && tbc.onChartRecommendedClick();
                    break;
                case 'pagesetup':
                    var lm = this.ctrl('LeftMenu');
                    lm && lm.leftMenu && lm.leftMenu.showMenu('file:printpreview');
                    break;
            }
        },

        // ---- ribbon behaviour (like DE.Controllers.RibbonTabs) -----------------------------
        setupRibbonBehaviour: function (tb) {
            var me = this, v = this.view;
            Common.UI.RibbonCollapse.install(tb, {
                labelOf: function (key) { return tb.groupLabelText ? tb.groupLabelText(key) : ''; }
            });
            var $box = tb.$el.find('.box-controls').first();
            if (!$box.length) return;
            var $slot = $('<div class="elx-ribbon-pin"></div>').appendTo($box);
            v.btnRibbonPin = new Common.UI.Button({
                cls: 'btn-toolbar', iconCls: 'toolbar__icon btn-elx-ribbon-collapse',
                hint: v.txtCollapseRibbon, dataHint: '0', dataHintDirection: 'top', dataHintOffset: 'small'
            });
            v.btnRibbonPin.render($slot);
            var sync = function () {
                var folded = tb.isCompact();
                v.btnRibbonPin.setIconCls('toolbar__icon ' + (folded ? 'btn-elx-ribbon-pin' : 'btn-elx-ribbon-collapse'));
                v.btnRibbonPin.updateHint(folded ? v.txtPinRibbon : v.txtCollapseRibbon);
            };
            v.btnRibbonPin.on('click', function () { me.toggleRibbon(); });
            tb.on('view:compact', function () { setTimeout(sync, 0); });
            sync();
        },

        toggleRibbon: function () {
            var tc = this.ctrl('Toolbar');
            tc && tc.onChangeViewMode && tc.onChangeViewMode(null, !this.toolbar.isCompact());
            this.view.btnRibbonPin && setTimeout(_.bind(function () {
                var folded = this.toolbar.isCompact(), v = this.view;
                v.btnRibbonPin.setIconCls('toolbar__icon ' + (folded ? 'btn-elx-ribbon-pin' : 'btn-elx-ribbon-collapse'));
                v.btnRibbonPin.updateHint(folded ? v.txtPinRibbon : v.txtCollapseRibbon);
            }, this), 0);
        },

        // Ctrl+F1 shows or hides the ribbon, like Excel
        bindRibbonShortcuts: function () {
            var me = this, keymap = {};
            keymap['ctrl+f1'] = function (e) {
                if (Common.Utils.ModalWindow && Common.Utils.ModalWindow.isVisible()) return true;
                me.toggleRibbon();
                e && e.preventDefault && e.preventDefault();
                return false;
            };
            Common.util.Shortcuts.delegateShortcuts({shortcuts: keymap});
        }
    }, SSE.Controllers.RibbonTabs || {}));
});

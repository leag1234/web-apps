/*
 * SPDX-FileCopyrightText: 2026 ElixirOffice contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 *
 * DE.Controllers.RibbonTabs: Word-like Insert, Layout, Design, References,
 * Review and View tabs (see DE.Views.RibbonTabs).
 *
 * Toolbar 'render:before': the template (toolbar.$layout) is rearranged in the
 * order of Word and the new controls are rendered into new slots, before the
 * stock controls are injected. Stock controls that Word shows differently are
 * kept, hidden (class elx-source), so that their controllers and locks keep
 * working; visible twins (Common.UI.RibbonMirror) follow their state.
 */

define([
    'core',
    'documenteditor/main/app/view/RibbonTabs'
], function () {
    'use strict';

    // Word's order of the shape categories; values: index in Main.shapeGroupNames
    var SHAPE_ORDER = [8, 7, 0, 1, 2, 3, 4, 5, 6],
        SHAPE_NAMES = ['txtShapesBasic', 'txtShapesBlockArrows', 'txtShapesEquation', 'txtShapesFlowchart', 'txtShapesStars',
                       'txtShapesCallouts', 'txtShapesButtons', 'txtShapesRectangles', 'txtShapesLines'];

    DE.Controllers.RibbonTabs = Backbone.Controller.extend(_.extend({
        models: [],
        collections: [],
        views: ['RibbonTabs'],

        initialize: function () {
            var me = this;
            this.addListeners({
                'Toolbar': {
                    'render:before': function (toolbar) { me.onToolbarBeforeRender(toolbar); },
                    'render:after': function (toolbar) { me.onToolbarAfterRender(toolbar); }
                }
            });
            Common.NotificationCenter.on('app:ready', this.onAppReady.bind(this));
        },

        onLaunch: function () {
            this.sortShapeGroups();
        },

        appOptions: function () {
            var main = this.getApplication().getController('Main');
            return (main && main.appOptions) || {};
        },

        onToolbarBeforeRender: function (toolbar) {
            var mode = this.appOptions();
            if (!mode.isEdit || mode.isEditMailMerge || mode.isEditDiagram || mode.isEditOle || this.view) return;
            this.toolbar = toolbar;
            this.view = this.createView('RibbonTabs', {toolbar: toolbar, mode: mode});
            this.layoutInsert(toolbar.$layout);
            this.layoutLayout(toolbar);
            this.view.createInsert();
            this.view.createLayout();
        },

        onToolbarAfterRender: function (toolbar) {
            if (!this.view) return;
            Array.prototype.push.apply(toolbar.lockControls, this.view.lockedControls);
            Common.Utils.lockControls(Common.enumLock.disableOnStart, true, {array: this.view.lockedControls});
            Common.Utils.lockControls('elx-no-object', true, {array: [this.view.btnLayPosition, this.view.btnLayRotate]});
            this.bindInsert();
            this.bindLayout();
            var api = this.getApplication().getController('Viewport').getApi();
            api && api.asc_registerCallback('asc_onFocusObject', _.bind(this.onApiFocusObject, this));
            this.api = api;
        },

        onAppReady: function () {
            if (!this.view) return;
            var me = this, v = this.view;
            Common.Utils.lockControls(Common.enumLock.disableOnStart, false, {array: v.lockedControls});
            // the sources of some twins are created after the toolbar (Links, Header & Footer)
            setTimeout(function () {
                v.attachTwins();
                me.renameInsertControls();
                me.renameLayoutControls();
            }, 0);
        },

        // a stock slot stays in the layout, hidden, so that its control and handlers still exist
        hideSource: function ($slot) {
            $slot.addClass('elx-source');
        },

        // ---- Insert tab ------------------------------------------------------------------
        layoutInsert: function ($layout) {
            var $ins = $layout.find('.panel[data-tab=ins]'),
                slot = function (id) { return $('<span class="btn-slot text x-huge" id="' + id + '"></span>'); };
            if (!$ins.length) return;

            // Pages: Blank Page, Page Break
            var $pb = $ins.find('.group[data-group=pages] .btn-pagebreak');
            $pb.after(slot('slot-elx-ins-pagebreak'));
            this.hideSource($pb);

            // Links: Link, Bookmark, Cross-reference
            var $link = $ins.find('.group[data-group=links] .slot-inshyperlink');
            $link.after(slot('slot-elx-ins-bookmark'), slot('slot-elx-ins-crossref'));

            // Header & Footer: Header, Footer, Page Number
            var $hf = $ins.find('.group[data-group=headerfooter] .slot-headerfooter');
            $hf.after(slot('slot-elx-ins-header'), slot('slot-elx-ins-footer'));
            this.hideSource($hf);

            // Text: Text Box, Quick Parts, Text Art, Drop Cap, Date & Time, Object
            var $text = $ins.find('.group[data-group=text]'),
                $field = $text.find('.slot-insertfield'),
                $file = $text.find('#slot-btn-text-from-file');
            $text.find('#slot-btn-instext').after(slot('slot-elx-ins-quickparts'));
            $text.find('.slot-insertdatetime').after(slot('slot-elx-ins-object'));
            this.hideSource($field);
            this.hideSource($file);
            // layout customization (customization.layout) hides them by these names
            $text.find('#slot-elx-ins-quickparts').attr('data-layout-name', 'toolbar-insert-field');
            $text.find('#slot-elx-ins-object').attr('data-layout-name', 'toolbar-insert-file');
        },

        bindInsert: function () {
            var me = this, v = this.view;
            var hf = function (menu, item) {
                var c = me.getApplication().getController('HeaderFooterTab');
                c && c.onHeaderFooterEditRemove(item);
            };
            v.btnInsHeader.menu.on('item:click', hf);
            v.btnInsFooter.menu.on('item:click', hf);
            v.btnInsQuickParts.menu.on('item:click', function (menu, item) {
                var src = v.btnInsQuickParts.mirrorSource();
                src && !src.isDisabled() && src.trigger('click', src);
            });
            v.btnInsObject.menu.on('item:click', function (menu, item) { v.btnInsObject.clickSourceItem(item.value); });
            v.btnInsObject.menu.on('show:before', function (menu) {
                var src = v.btnInsObject.mirrorSource();
                if (!src || !src.menu || !_.isObject(src.menu)) return;
                _.each(menu.items, function (it) {
                    var si = _.find(src.menu.items, function (s) { return s.value === it.value; });
                    si && it.setVisible(si.$el ? si.$el.css('display') !== 'none' : true);
                    si && it.setDisabled(si.isDisabled());
                });
            });
        },

        renameInsertControls: function () {
            var v = this.view, c = this.getApplication().getController('HeaderFooterTab'),
                hv = c && c.getView('HeaderFooterTab');
            // Word: "Pictures", "Shapes"
            hv && _.each(hv.btnsInsImage || [], function (b) {
                b.cmpEl && b.cmpEl.closest('.panel[data-tab=ins]').length && b.setCaption(v.txtPictures);
            });
            this.toolbar.btnInsertShape && this.toolbar.btnInsertShape.setCaption(v.txtShapes);
            // Table menu items as in Word
            var mt = this.toolbar.mnuInsertTable;
            if (mt) {
                _.each(mt.items, function (it) {
                    it.value === 'custom' && it.setCaption(v.txtInsertTableDlg);
                    it.value === 'convert' && it.setCaption(v.txtConvertTextToTable);
                });
            }
            this.toolbar.processPanelVisible && this.toolbar.processPanelVisible(null, true);
        },

        // ---- Layout tab ------------------------------------------------------------------
        layoutLayout: function (tb) {
            var $lay = tb.$layout.find('.panel[data-tab=layout]');
            if (!$lay.length) return;
            var slot = function (id, cls) { return $('<span class="btn-slot text' + (cls === undefined ? ' x-huge' : cls) + '" id="' + id + '"></span>'); };

            // Page Setup: Margins, Orientation, Size, Columns + a column of Breaks, Line Numbers, Hyphenation
            var $ps = $lay.find('.group[data-group=pagesetup]'),
                $margins = $ps.find('#slot-btn-pagemargins');
            $margins.after(slot('slot-elx-lay-margins'));
            this.hideSource($margins);
            var $pb = $ps.find('.btn-pagebreak');
            this.hideSource($pb);
            var $ln = $ps.find('#slot-btn-line-numbers').removeClass('x-huge'),
                $hy = $ps.find('#slot-btn-hyphenation').removeClass('x-huge'),
                $col = $('<div class="elset-column elx-rows3"></div>');
            $col.append($('<div class="elset"></div>').append(slot('slot-elx-lay-breaks', '')),
                        $('<div class="elset"></div>').append($ln),
                        $('<div class="elset"></div>').append($hy));
            $ps.append($col);
            _.each([tb.btnLineNumbers, tb.btnHyphenation], function (b) {
                if (b) { b.cls = b.options.cls = 'btn-toolbar'; b.options.dataHintOffset = '0, -6'; }
            });

            // Paragraph: "Indent" and "Spacing" headings, like Word
            var $pgs = $lay.find('#slot-spin-ind-left').closest('.group'),
                $sps = $lay.find('#slot-spin-space-before').closest('.group');
            $pgs.addClass('elx-rows3').prepend('<div class="elset elx-subhead">' + this.view.txtIndent + '</div>');
            $sps.addClass('elx-rows3').prepend('<div class="elset elx-subhead margin-left-6">' + this.view.txtSpacing + '</div>');

            // Arrange: Position, Wrap Text, Bring Forward, Send Backward, Align, Group, Rotate
            $lay.find('#slot-img-wrapping').before(slot('slot-elx-lay-position'));
            $lay.find('#slot-img-group').after(slot('slot-elx-lay-rotate'));
        },

        bindLayout: function () {
            var me = this, v = this.view;
            var ctx = function () { return me.getApplication().getController('ContextTabs'); };
            // Margins gallery
            v.btnLayMargins.menu.on('show:before', function (menu) { me.updateMarginsMenu(menu); });
            v.btnLayMargins.menu.on('item:click', function (menu, item) { me.onMarginsClick(item); });
            // Breaks
            v.btnLayBreaks.menu.on('item:click', function (menu, item) {
                item.value !== undefined && me.toolbar.fireEvent('insert:break', [item.value]);
            });
            // Position and Rotate share the Picture/Shape Format handlers
            v.btnLayPosition.menu.on('show:before', function (menu) {
                if (menu.items.length) return;
                var cv = ctx() && ctx().view;
                cv && _.each(cv._positionItems(), function (it) { menu.addItem(it.caption === '--' ? {caption: '--'} : new Common.UI.MenuItem(it)); });
            });
            v.btnLayRotate.menu.on('show:before', function (menu) {
                if (menu.items.length) return;
                var cv = ctx() && ctx().view;
                cv && _.each(cv.picArrange.rotate.menu.items, function (it) {
                    menu.addItem(it.caption === '--' ? {caption: '--'} : new Common.UI.MenuItem({caption: it.caption, iconCls: it.options.iconCls, value: it.value}));
                });
            });
            v.btnLayPosition.menu.on('item:click', function (menu, item) { ctx() && ctx().onPositionClick(item); });
            v.btnLayRotate.menu.on('item:click', function (menu, item) { ctx() && ctx().onRotateClick(item); });
        },

        renameLayoutControls: function () {
            var tb = this.toolbar, v = this.view;
            tb.btnImgWrapping && tb.btnImgWrapping.setCaption(v.txtWrapText);
            tb.lblIndentsLeft && tb.lblIndentsLeft.setCaption(v.txtIndLeft);
            tb.lblIndentsRight && tb.lblIndentsRight.setCaption(v.txtIndRight);
            tb.lblSpacingBefore && tb.lblSpacingBefore.setCaption(v.txtBefore);
            tb.lblSpacingAfter && tb.lblSpacingAfter.setCaption(v.txtAfter);
            tb.processPanelVisible && tb.processPanelVisible(null, true);
        },

        onApiFocusObject: function (selected) {
            if (!this.view) return;
            var obj = null;
            _.each(selected, function (s) {
                if (s.get_ObjectType() === Asc.c_oAscTypeSelectElement.Image && !s.get_ObjectValue().get_ChartProperties()) obj = s.get_ObjectValue();
            });
            var dis = !obj || obj.get_Locked();
            Common.Utils.lockControls('elx-no-object', dis, {array: [this.view.btnLayRotate]});
            Common.Utils.lockControls('elx-no-object', dis || !obj.get_CanBeFlow() || obj.get_FromGroup(), {array: [this.view.btnLayPosition]});
        },

        updateMarginsMenu: function (menu) {
            if (!this.api) return;
            var sp = this.api.asc_GetSectionProps(),
                cur = [sp.get_TopMargin(), sp.get_LeftMargin(), sp.get_BottomMargin(), sp.get_RightMargin()],
                mirror = !!(sp.get_MirrorMargins && sp.get_MirrorMargins()),
                ls = Common.localStorage,
                last = ls.itemExists('de-pgmargins-top') ? [parseFloat(ls.getItem('de-pgmargins-top')), parseFloat(ls.getItem('de-pgmargins-left')),
                    parseFloat(ls.getItem('de-pgmargins-bottom')), parseFloat(ls.getItem('de-pgmargins-right'))] : null;
            var same = function (a) { return a && _.every(a, function (x, i) { return Math.abs(x - cur[i]) < 0.1; }); };
            _.each(menu.items, function (it) {
                if (it.options.elxId === 'last') {
                    it.options.value = it.value = last;
                    it.setVisible(!!last);
                }
                if (it.options.checkable) it.setChecked(!!it.value && same(it.value) && !!it.options.mirror === mirror, true);
            });
        },

        onMarginsClick: function (item) {
            var tc = this.getApplication().getController('Toolbar'), api = this.api;
            if (!api || !item) return;
            if (item.value === 'advanced') { tc.onPageMarginsSelect(null, {value: 'advanced'}); return; }
            var m = item.value;
            if (!m) return;
            if (tc.checkPageSize && tc.checkPageSize(undefined, undefined, m[1], m[3], m[0], m[2])) {
                tc.onSectionProps(api.asc_GetSectionProps());
                return;
            }
            var props = new Asc.CDocumentSectionProps();
            props.put_TopMargin(m[0]);
            props.put_LeftMargin(m[1]);
            props.put_BottomMargin(m[2]);
            props.put_RightMargin(m[3]);
            props.put_MirrorMargins && props.put_MirrorMargins(!!item.options.mirror);
            api.asc_SetSectionProps(props);
            Common.NotificationCenter.trigger('edit:complete', this.toolbar);
        },

        // Shapes gallery: Word's categories and order
        sortShapeGroups: function () {
            var me = this, store = this.getApplication().getCollection('ShapeGroups'),
                main = this.getApplication().getController('Main');
            if (!store || !main) return;
            var rank = function (model) {
                var idx = main.shapeGroupNames ? main.shapeGroupNames.indexOf(model.get('elxGroupName') || model.get('groupName')) : -1;
                var r = SHAPE_ORDER.indexOf(idx);
                return r < 0 ? 99 : r;
            };
            store.comparator = function (a, b) { return rank(a) - rank(b); };
            store.on('add', function (model) {
                if (model.get('elxGroupName')) return;
                var idx = main.shapeGroupNames ? main.shapeGroupNames.indexOf(model.get('groupName')) : -1;
                model.set('elxGroupName', model.get('groupName'), {silent: true});
                var v = me.view || DE.Views.RibbonTabs.prototype;
                idx >= 0 && v[SHAPE_NAMES[idx]] && model.set('groupName', v[SHAPE_NAMES[idx]], {silent: true});
            });
        }
    }, DE.Controllers.RibbonTabs || {}));
});

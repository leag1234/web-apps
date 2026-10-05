/*
 * SPDX-FileCopyrightText: 2026 ElixirOffice contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 *
 * DE.Views.RibbonTabs: Word-like content of the Insert, Layout, Design,
 * References, Review and View tabs.
 *
 * The ribbon template (Toolbar.template) and the controls of the other views are
 * left as they are: DE.Controllers.RibbonTabs reorders the template right before
 * the controls are injected (Toolbar 'render:before') and adds the controls
 * created here. Commands that exist elsewhere are mirrored (Common.UI.RibbonMirror)
 * so their controllers, locks and shortcuts are unchanged.
 */

define([
    'core',
    'common/main/lib/component/Button',
    'common/main/lib/component/RibbonMirror'
], function () {
    'use strict';

    DE.Views.RibbonTabs = Common.UI.BaseView.extend(_.extend((function () {
        return {
            options: {},

            initialize: function (options) {
                Common.UI.BaseView.prototype.initialize.call(this);
                this.toolbar = options.toolbar;
                this.mode = options.mode;
                this.lockedControls = [];
                this.twins = [];
            },

            _baseLock: function () {
                var _set = Common.enumLock;
                return [_set.previewReviewMode, _set.viewFormMode, _set.lostConnect, _set.disableOnStart, _set.docLockView, _set.docLockForms, _set.docLockComments, _set.viewMode];
            },

            // a big ribbon button rendered into slot #id of the toolbar layout
            btn: function (id, cfg) {
                var b = new Common.UI.Button(_.extend({
                    cls: 'btn-toolbar x-huge icon-top',
                    lock: this._baseLock(),
                    dataHint: '1',
                    dataHintDirection: 'bottom',
                    dataHintOffset: 'small'
                }, cfg));
                b.render(this.toolbar.$layout.find('#' + id));
                this.lockedControls.push(b);
                return b;
            },

            mirror: function (id, src, opts) {
                // the source may belong to a view that does not exist yet
                var safe = _.isFunction(src) ? function () { try { return src(); } catch (e) { return undefined; } } : src;
                var twin = Common.UI.RibbonMirror(safe, opts);
                twin.render(this.toolbar.$layout.find('#' + id));
                this.twins.push(twin);
                return twin;
            },

            attachTwins: function () {
                _.each(this.twins, function (t) { t.mirrorAttach(); });
            },

            // ---- Insert ---------------------------------------------------------------
            createInsert: function () {
                var me = this, tb = this.toolbar,
                    app = DE.getController('HeaderFooterTab'),
                    hfView = function () { return app && app.getView('HeaderFooterTab'); },
                    insSlot = function (list) {
                        return _.find(list || [], function (b) { return b.cmpEl && b.cmpEl.closest('.panel[data-tab=ins]').length; });
                    };

                // Pages: Page Break (Word), the Breaks menu lives in Layout
                this.btnInsPageBreak = this.mirror('slot-elx-ins-pagebreak', function () { return tb.btnsPageBreak && tb.btnsPageBreak[0]; }, {
                    caption: this.txtPageBreak, iconCls: 'toolbar__icon btn-pagebreak', split: false,
                    hint: this.tipPageBreak,
                    click: function () { tb.fireEvent('insert:break', ['page']); }
                });

                // Links: Bookmark and Cross-reference, like Word
                var links = function () { var c = DE.getController('Links'); return c && c.getView('Links'); };
                this.btnInsBookmark = this.mirror('slot-elx-ins-bookmark', function () { var v = links(); return v && v.btnBookmarks; },
                    {caption: this.txtBookmark, iconCls: 'toolbar__icon btn-bookmarks'});
                this.btnInsCrossRef = this.mirror('slot-elx-ins-crossref', function () { var v = links(); return v && v.btnCrossRef; },
                    {caption: this.txtCrossReference, iconCls: 'toolbar__icon btn-cross-reference'});

                // Header & Footer: separate Header and Footer buttons
                var hfSrc = function () { var v = hfView(); return v && insSlot(v.btnsHeaderFooter); };
                this.btnInsHeader = this.mirror('slot-elx-ins-header', hfSrc, {
                    caption: this.txtHeader, iconCls: 'toolbar__icon btn-elx-header', split: false,
                    menu: new Common.UI.Menu({items: [
                        {caption: this.txtEditHeader, value: 'edit-header'},
                        {caption: this.txtRemoveHeader, value: 'remove-header'}
                    ]})
                });
                this.btnInsFooter = this.mirror('slot-elx-ins-footer', hfSrc, {
                    caption: this.txtFooter, iconCls: 'toolbar__icon btn-elx-footer', split: false,
                    menu: new Common.UI.Menu({items: [
                        {caption: this.txtEditFooter, value: 'edit-footer'},
                        {caption: this.txtRemoveFooter, value: 'remove-footer'}
                    ]})
                });

                // Text: Quick Parts (fields) and Object (text from file), like Word
                var fieldSrc = function () { var v = hfView(); return v && insSlot(v.btnsInsField); };
                this.btnInsQuickParts = this.mirror('slot-elx-ins-quickparts', fieldSrc, {
                    caption: this.txtQuickParts, iconCls: 'toolbar__icon btn-elx-quick-parts', split: false,
                    menu: new Common.UI.Menu({items: [{caption: this.txtField, value: 'field'}]})
                });
                this.btnInsObject = this.mirror('slot-elx-ins-object', function () { return tb.btnTextFromFile; }, {
                    caption: this.txtObject, iconCls: 'toolbar__icon btn-text-from-file', split: false,
                    menu: new Common.UI.Menu({items: [
                        {caption: this.txtTextFromFile, value: 'file'},
                        {caption: this.txtTextFromUrl, value: 'url'},
                        {caption: this.txtTextFromStorage, value: 'storage'}
                    ]})
                });
            },

            // ---- strings ----------------------------------------------------------------
            txtPageBreak: 'Page Break',
            tipPageBreak: 'Start the next page at the current position',
            txtBookmark: 'Bookmark',
            txtCrossReference: 'Cross-reference',
            txtHeader: 'Header',
            txtFooter: 'Footer',
            txtEditHeader: 'Edit Header',
            txtRemoveHeader: 'Remove Header',
            txtEditFooter: 'Edit Footer',
            txtRemoveFooter: 'Remove Footer',
            txtQuickParts: 'Quick Parts',
            txtField: 'Field...',
            txtObject: 'Object',
            txtTextFromFile: 'Text from File...',
            txtTextFromUrl: 'Text from URL...',
            txtTextFromStorage: 'Text from Storage...',
            txtPictures: 'Pictures',
            txtShapes: 'Shapes',
            txtInsertTableDlg: 'Insert Table...',
            txtConvertTextToTable: 'Convert Text to Table...',
            txtShapesLines: 'Lines',
            txtShapesRectangles: 'Rectangles',
            txtShapesBasic: 'Basic Shapes',
            txtShapesBlockArrows: 'Block Arrows',
            txtShapesEquation: 'Equation Shapes',
            txtShapesFlowchart: 'Flowchart',
            txtShapesStars: 'Stars and Banners',
            txtShapesCallouts: 'Callouts',
            txtShapesButtons: 'Action Buttons'
        };
    }()), DE.Views.RibbonTabs || {}));
});

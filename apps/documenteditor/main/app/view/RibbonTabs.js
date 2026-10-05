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

            // ---- Layout -------------------------------------------------------------------
            // Word's margin presets, [top, left, bottom, right] in mm; mirrored: inside/outside
            marginPresets: function () {
                var us = /^(ca|us)$/i.test(Common.Utils.InternalSettings.get('de-config-region'));
                return [
                    {id: 'normal', caption: this.txtMarginsNormal, value: us ? [25.4, 25.4, 25.4, 25.4] : [25, 25, 25, 25]},
                    {id: 'narrow', caption: this.txtMarginsNarrow, value: [12.7, 12.7, 12.7, 12.7]},
                    {id: 'moderate', caption: this.txtMarginsModerate, value: [25.4, 19.05, 25.4, 19.05]},
                    {id: 'wide', caption: this.txtMarginsWide, value: [25.4, 50.8, 25.4, 50.8]},
                    {id: 'mirrored', caption: this.txtMarginsMirrored, value: [25.4, 31.75, 25.4, 25.4], mirror: true},
                    {id: 'office2003', caption: this.txtMarginsOffice2003, value: [25.4, 31.75, 25.4, 31.75]}
                ];
            },

            marginItemTemplate: function () {
                var me = this;
                return _.template('<a id="<%= id %>" tabindex="-1" type="menuitem" class="elx-gallery-item">' +
                    '<span class="elx-gallery-icon"><svg class="icon"><use href="#<%= options.icon %>"></use></svg></span>' +
                    '<span class="elx-gallery-text"><b><%= caption %></b>' +
                    '<% if (options.value) { var m = Common.Utils.Metric, f = function (v) { return parseFloat(m.fnRecalcFromMM(v).toFixed(2)) + " " + m.getCurrentMetricName(); }; %>' +
                    '<span class="elx-gallery-grid">' +
                        '<span>' + me.txtTop + ' <%= f(options.value[0]) %></span><span>' + me.txtBottom + ' <%= f(options.value[2]) %></span>' +
                        '<span><%= options.mirror ? "' + me.txtInside + '" : "' + me.txtLeft + '" %> <%= f(options.value[1]) %></span>' +
                        '<span><%= options.mirror ? "' + me.txtOutside + '" : "' + me.txtRight + '" %> <%= f(options.value[3]) %></span>' +
                    '</span><% } %></span></a>');
            },

            breakItemTemplate: function () {
                return _.template('<a id="<%= id %>" tabindex="-1" type="menuitem" class="elx-gallery-item">' +
                    '<span class="elx-gallery-icon"><svg class="icon"><use href="#<%= options.icon %>"></use></svg></span>' +
                    '<span class="elx-gallery-text"><b><%= caption %></b><span class="elx-gallery-desc"><%= options.desc %></span></span></a>');
            },

            createLayout: function () {
                var me = this, tb = this.toolbar;
                var mi = this.marginItemTemplate();
                var ls = Common.localStorage,
                    last = ls.itemExists('de-pgmargins-top') ? _.map(['top', 'left', 'bottom', 'right'], function (k) { return parseFloat(ls.getItem('de-pgmargins-' + k)); }) : null;
                var items = [{caption: this.txtMarginsLast, value: last, elxId: 'last', icon: 'btn-elx-margins-last', template: mi, checkable: true, toggleGroup: 'elx-margins'}];
                _.each(this.marginPresets(), function (p) {
                    items.push({caption: p.caption, value: p.value, mirror: !!p.mirror, preset: p.id, icon: 'btn-elx-margins-' + p.id, template: mi, checkable: true, toggleGroup: 'elx-margins'});
                });
                items.push({caption: '--'}, {caption: this.txtCustomMargins, value: 'advanced'});
                this.btnLayMargins = this.mirror('slot-elx-lay-margins', function () { return tb.btnPageMargins; }, {
                    caption: tb.capBtnMargins, iconCls: 'toolbar__icon btn-pagemargins', split: false,
                    menu: new Common.UI.Menu({cls: 'elx-gallery-menu', items: items})
                });

                var bi = this.breakItemTemplate(), hdr = function (t) { return {template: _.template('<div class="elx-menu-header">' + t + '</div>')}; };
                this.btnLayBreaks = this.mirror('slot-elx-lay-breaks', function () { return tb.btnsPageBreak && tb.btnsPageBreak[0]; }, {
                    cls: 'btn-toolbar', caption: this.txtBreaks, iconCls: 'toolbar__icon btn-pagebreak', split: false,
                    menu: new Common.UI.Menu({cls: 'elx-gallery-menu', items: [
                        hdr(this.txtPageBreaks),
                        {caption: this.txtBreakPage, desc: this.txtBreakPageDesc, value: 'page', icon: 'btn-elx-brk-page', template: bi},
                        {caption: this.txtBreakColumn, desc: this.txtBreakColumnDesc, value: 'column', icon: 'btn-elx-brk-column', template: bi},
                        hdr(this.txtSectionBreaks),
                        {caption: this.txtBreakNext, desc: this.txtBreakNextDesc, value: Asc.c_oAscSectionBreakType.NextPage, icon: 'btn-elx-brk-next', template: bi},
                        {caption: this.txtBreakCont, desc: this.txtBreakContDesc, value: Asc.c_oAscSectionBreakType.Continuous, icon: 'btn-elx-brk-cont', template: bi},
                        {caption: this.txtBreakEven, desc: this.txtBreakEvenDesc, value: Asc.c_oAscSectionBreakType.EvenPage, icon: 'btn-elx-brk-even', template: bi},
                        {caption: this.txtBreakOdd, desc: this.txtBreakOddDesc, value: Asc.c_oAscSectionBreakType.OddPage, icon: 'btn-elx-brk-odd', template: bi}
                    ]})
                });

                // Arrange: Position and Rotate (handlers of DE.Controllers.ContextTabs)
                var objLock = this._baseLock().concat(['elx-no-object']);
                this.btnLayPosition = this.btn('slot-elx-lay-position', {
                    iconCls: 'toolbar__icon btn-elx-position', caption: this.txtPosition, lock: objLock,
                    menu: new Common.UI.Menu({cls: 'elx-position-menu', items: []})
                });
                this.btnLayRotate = this.btn('slot-elx-lay-rotate', {
                    iconCls: 'toolbar__icon btn-elx-rotate', caption: this.txtRotate, lock: objLock,
                    menu: new Common.UI.Menu({items: []})
                });
            },

            // a control rendered into an existing (live) slot
            liveBtn: function ($slot, cfg) {
                var b = new Common.UI.Button(_.extend({
                    cls: 'btn-toolbar x-huge icon-top', lock: this._baseLock(),
                    dataHint: '1', dataHintDirection: 'bottom', dataHintOffset: 'small'
                }, cfg));
                b.render($slot);
                this.lockedControls.push(b);
                return b;
            },

            liveMirror: function ($slot, src, opts) {
                var safe = _.isFunction(src) ? function () { try { return src(); } catch (e) { return undefined; } } : src;
                var twin = Common.UI.RibbonMirror(safe, opts);
                twin.render($slot);
                this.twins.push(twin);
                twin.mirrorAttach();
                return twin;
            },

            // Word's watermark presets
            watermarkPresets: function () {
                return [this.txtWmConfidential, this.txtWmDoNotCopy, this.txtWmDraft, this.txtWmSample, this.txtWmUrgent, this.txtWmAsap];
            },

            // ---- strings ----------------------------------------------------------------
            txtWordCount: 'Word Count',
            txtWordCountTitle: 'Word Count',
            txtStatistics: 'Statistics:',
            txtPages: 'Pages',
            txtWords: 'Words',
            txtCharsNoSpaces: 'Characters (no spaces)',
            txtCharsSpaces: 'Characters (with spaces)',
            txtParagraphs: 'Paragraphs',
            txtCounting: 'Counting...',
            txtSpelling: 'Spelling',
            tipSpelling: 'Check spelling as you type',
            txtLanguage: 'Language',
            txtSetProofingLanguage: 'Set Proofing Language...',
            txtPrevComment: 'Previous',
            txtNextComment: 'Next',
            tipPrevComment: 'Go to the previous comment',
            tipNextComment: 'Go to the next comment',
            txtShowComments: 'Show Comments',
            txtTrackChanges: 'Track Changes',
            txtForEveryone: 'For Everyone',
            txtJustMine: 'Just Mine',
            txtAllMarkup: 'All Markup',
            txtSimpleMarkup: 'Simple Markup',
            txtNoMarkup: 'No Markup',
            txtOriginal: 'Original',
            txtPrevChange: 'Previous',
            txtNextChange: 'Next',
            txtInsertFootnote: 'Insert Footnote',
            txtInsertEndnote: 'Insert Endnote',
            txtNextFootnote: 'Next Footnote',
            txtPrevFootnote: 'Previous Footnote',
            txtNextEndnote: 'Next Endnote',
            txtPrevEndnote: 'Previous Endnote',
            txtInsertCaption: 'Insert Caption',
            txtReadMode: 'Read Mode',
            tipReadMode: 'Read the document without editing it',
            txtPrintLayout: 'Print Layout',
            tipPrintLayout: 'Edit the document as it will be printed',
            txtRuler: 'Ruler',
            txtNavigationPane: 'Navigation Pane',
            txtOnePage: 'One Page',
            txtPageWidth: 'Page Width',
            txtMultiplePages: 'Multiple Pages',
            txtZoom100: '100%',
            txtWatermark: 'Watermark',
            txtCustomWatermark: 'Custom Watermark...',
            txtRemoveWatermark: 'Remove Watermark',
            txtWmConfidential: 'CONFIDENTIAL',
            txtWmDoNotCopy: 'DO NOT COPY',
            txtWmDraft: 'DRAFT',
            txtWmSample: 'SAMPLE',
            txtWmUrgent: 'URGENT',
            txtWmAsap: 'ASAP',
            txtGroupProofing: 'Proofing',
            txtGroupLanguage: 'Language',
            txtGroupChanges: 'Changes',
            txtGroupViews: 'Views',
            txtGroupWindow: 'Interface',
            txtMarginsLast: 'Last Custom Setting',
            txtMarginsNormal: 'Normal',
            txtMarginsNarrow: 'Narrow',
            txtMarginsModerate: 'Moderate',
            txtMarginsWide: 'Wide',
            txtMarginsMirrored: 'Mirrored',
            txtMarginsOffice2003: 'Office 2003 Default',
            txtCustomMargins: 'Custom Margins...',
            txtTop: 'Top:',
            txtBottom: 'Bottom:',
            txtLeft: 'Left:',
            txtRight: 'Right:',
            txtInside: 'Inside:',
            txtOutside: 'Outside:',
            txtBreaks: 'Breaks',
            txtPageBreaks: 'Page Breaks',
            txtSectionBreaks: 'Section Breaks',
            txtBreakPage: 'Page',
            txtBreakPageDesc: 'Mark the point at which one page ends and the next page begins.',
            txtBreakColumn: 'Column',
            txtBreakColumnDesc: 'Indicate that the text following the column break will begin in the next column.',
            txtBreakNext: 'Next Page',
            txtBreakNextDesc: 'Insert a section break and start the new section on the next page.',
            txtBreakCont: 'Continuous',
            txtBreakContDesc: 'Insert a section break and start the new section on the same page.',
            txtBreakEven: 'Even Page',
            txtBreakEvenDesc: 'Insert a section break and start the new section on the next even-numbered page.',
            txtBreakOdd: 'Odd Page',
            txtBreakOddDesc: 'Insert a section break and start the new section on the next odd-numbered page.',
            txtPosition: 'Position',
            txtRotate: 'Rotate',
            txtWrapText: 'Wrap Text',
            txtIndent: 'Indent',
            txtSpacing: 'Spacing',
            txtIndLeft: 'Left:',
            txtIndRight: 'Right:',
            txtBefore: 'Before:',
            txtAfter: 'After:',
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

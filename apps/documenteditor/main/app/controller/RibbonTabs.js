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
            // group labels of the new groups (see Toolbar.updateGroupLabels)
            var v = this.view;
            _.each(_.keys(v.constructor.prototype), function (k) {
                if (/^txtGroup/.test(k) && toolbar[k] === undefined) toolbar[k] = v[k];
            });
            this.layoutInsert(toolbar.$layout);
            this.layoutLayout(toolbar);
            this.layoutReferences(toolbar.$layout);
            this.layoutDesign(toolbar.$layout);
            this.view.createInsert();
            this.view.createLayout();
            this.createReferences();
            this.createDesign();
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
            // Review and View are built by their own controllers right after the toolbar
            var me = this;
            setTimeout(function () {
                me.arrangeReview();
                me.arrangeView();
                Array.prototype.push.apply(toolbar.lockControls, me.view.lockedControls);
                !me._docReady && Common.Utils.lockControls(Common.enumLock.disableOnStart, true, {array: me.view.lockedControls});
            }, 0);
        },

        onAppReady: function () {
            if (!this.view) return;
            var me = this, v = this.view;
            this._docReady = true;
            Common.Utils.lockControls(Common.enumLock.disableOnStart, false, {array: v.lockedControls});
            // the sources of some twins are created after the toolbar (Links, Header & Footer)
            setTimeout(function () {
                v.attachTwins();
                me.renameInsertControls();
                me.renameLayoutControls();
                me.renameReviewControls();
                me.renameViewControls();
                me.renameReferencesControls();
                me.toolbar.processPanelVisible && me.toolbar.processPanelVisible(null, true);
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

        // ---- References -------------------------------------------------------------------
        layoutReferences: function ($layout) {
            var $ref = $layout.find('.panel[data-tab=links]');
            if (!$ref.length) return;
            // Footnotes: Insert Footnote + a column of Insert Endnote, Next Footnote
            var $fn = $ref.find('.group[data-group=footnotes]'), $notes = $fn.find('.slot-notes');
            this.hideSource($notes);
            $fn.attr('data-launcher', 'footnotes');
            $fn.append('<span class="btn-slot text x-huge" id="slot-elx-ref-footnote"></span>' +
                '<div class="elset-column"><div class="elset"><span class="btn-slot text" id="slot-elx-ref-endnote"></span></div>' +
                '<div class="elset"><span class="btn-slot text" id="slot-elx-ref-nextnote"></span></div></div>');
            // Link and Bookmark live in Insert, like Word
            var $links = $ref.find('.group[data-group=links]');
            $links.addClass('elx-hidden-group');
            $links.prev('.separator').addClass('elx-hidden-group');
            // Captions: Insert Caption + Insert Table of Figures, Update Table, Cross-reference
            var $cap = $ref.find('.group[data-group=captions]'), $xref = $cap.find('#slot-btn-crossref');
            this.hideSource($xref);
            var $colc = $cap.find('.elset-column').addClass('elx-rows3');
            $colc.append('<div class="elset"><span class="btn-slot text" id="slot-elx-ref-crossref"></span></div>');
        },

        createReferences: function () {
            var me = this, v = this.view, $l = this.toolbar.$layout,
                links = function () { var c = DE.getController('Links'); return c && c.getView('Links'); },
                notes = function (value) { var c = me.getApplication().getController('Links'); c && c.onNotesClick(value); },
                notesSrc = function () { var lv = links(); return lv && lv.btnsNotes && lv.btnsNotes[0]; };
            v.btnRefFootnote = v.liveMirror($l.find('#slot-elx-ref-footnote'), notesSrc, {
                caption: v.txtInsertFootnote, iconCls: 'toolbar__icon btn-notes', split: true,
                click: function () { notes('ins_footnote'); }
            });
            v.btnRefEndnote = v.liveMirror($l.find('#slot-elx-ref-endnote'), notesSrc, {
                cls: 'btn-toolbar', caption: v.txtInsertEndnote, iconCls: 'toolbar__icon btn-elx-endnote', split: false, menu: null,
                click: function () { notes('ins_endnote'); }
            });
            v.btnRefNextNote = v.liveMirror($l.find('#slot-elx-ref-nextnote'), notesSrc, {
                cls: 'btn-toolbar', caption: v.txtNextFootnote, iconCls: 'toolbar__icon btn-elx-next-note', split: true,
                menu: new Common.UI.Menu({items: [
                    {caption: v.txtNextFootnote, value: 'next'},
                    {caption: v.txtPrevFootnote, value: 'prev'},
                    {caption: v.txtNextEndnote, value: 'next-end'},
                    {caption: v.txtPrevEndnote, value: 'prev-end'}
                ]}),
                click: function () { notes('next'); }
            });
            v.btnRefNextNote.menu.on('item:click', function (menu, item) { notes(item.value); });
            v.btnRefCrossRef = v.liveMirror($l.find('#slot-elx-ref-crossref'), function () { var lv = links(); return lv && lv.btnCrossRef; }, {
                cls: 'btn-toolbar', caption: v.txtCrossReference, iconCls: 'toolbar__icon btn-elx-cross-ref-small'
            });
        },

        // ---- Design ---------------------------------------------------------------------------
        layoutDesign: function ($layout) {
            var $wm = $layout.find('.panel[data-tab=design] #slot-btn-watermark');
            if (!$wm.length) return;
            $wm.after('<span class="btn-slot text x-huge" id="slot-elx-design-watermark"></span>');
            this.hideSource($wm);
        },

        createDesign: function () {
            var me = this, v = this.view, tb = this.toolbar,
                items = _.map(v.watermarkPresets(), function (t, i) { return {caption: t, value: 'preset', text: t}; });
            items.push({caption: '--'}, {caption: v.txtCustomWatermark, value: 'edit'}, {caption: v.txtRemoveWatermark, value: 'remove'});
            v.btnDesignWatermark = v.liveMirror(tb.$layout.find('#slot-elx-design-watermark'), function () { return tb.btnWatermark; }, {
                caption: tb.capBtnWatermark, iconCls: 'toolbar__icon btn-watermark', split: false,
                menu: new Common.UI.Menu({cls: 'elx-watermark-menu', items: items})
            });
            v.btnDesignWatermark.menu.on('item:click', function (menu, item) {
                if (item.value === 'preset') me.applyWatermark(item.options.text);
                else v.btnDesignWatermark.clickSourceItem(item.value);
            });
        },

        // a Word watermark preset: diagonal semi-transparent grey text
        applyWatermark: function (text) {
            var api = this.api;
            if (!api) return;
            var props = api.asc_GetWatermarkProps();
            props.put_Type(Asc.c_oAscWatermarkType.Text);
            props.put_Text(text);
            props.put_IsDiagonal(true);
            props.put_Opacity(128);
            var tp = props.get_TextPr() || new Asc.CTextProp();
            tp.put_FontSize(-1);
            var font = new AscCommon.asc_CTextFontFamily();
            font.put_Name('Calibri');
            font.put_Index(-1);
            tp.put_FontFamily(font);
            tp.put_Bold(false);
            tp.put_Italic(false);
            var color = Common.Utils.ThemeColor.getRgbColor('C0C0C0');
            tp.put_Color(color);
            props.put_TextPr(tp);
            api.asc_SetWatermarkProps(props);
            Common.NotificationCenter.trigger('edit:complete', this.toolbar);
        },

        // ---- Review -----------------------------------------------------------------------------
        arrangeReview: function () {
            var me = this, v = this.view, tb = this.toolbar,
                rc = this.getApplication().getController('Common.Controllers.ReviewChanges'),
                rv = rc && rc.getView('Common.Views.ReviewChanges'),
                $rp = tb.$el.find('.panel[data-tab=review]');
            if (!$rp.length || !rv) return;
            var sep = function (cls) { return $('<div class="separator long' + (cls ? ' ' + cls : '') + '"></div>'); },
                grp = function (name, inner, cls) { return $('<div class="group' + (cls ? ' ' + cls : '') + '" data-group="' + name + '">' + inner + '</div>'); },
                big = function (id) { return '<span class="btn-slot text x-huge" id="' + id + '"></span>'; },
                small = function (id) { return '<div class="elset"><span class="btn-slot text" id="' + id + '"></span></div>'; };

            // Proofing and Language, first like Word
            var $proof = grp('proofing', big('slot-elx-rev-spelling') + big('slot-elx-rev-wordcount')),
                $lang = grp('language', big('slot-elx-rev-language'));
            $rp.prepend($proof, sep(), $lang, sep());
            var app = this.getApplication(),
                sbView = function () { var c = app.getController('Statusbar'); return c && c.getView('Statusbar'); };
            v.btnRevSpelling = v.liveMirror($rp.find('#slot-elx-rev-spelling'), function () { return rv.btnsSpelling && rv.btnsSpelling[0]; }, {
                caption: v.txtSpelling, iconCls: 'toolbar__icon btn-elx-spelling', hint: v.tipSpelling
            });
            v.btnRevWordCount = v.liveBtn($rp.find('#slot-elx-rev-wordcount'), {caption: v.txtWordCount, iconCls: 'toolbar__icon btn-elx-word-count'});
            v.btnRevWordCount.on('click', function () { me.showWordCount(); });
            v.btnRevLanguage = v.liveMirror($rp.find('#slot-elx-rev-language'), function () { return rv.btnsDocLang && rv.btnsDocLang[0]; }, {
                caption: v.txtLanguage, iconCls: 'toolbar__icon btn-elx-language', split: false,
                menu: new Common.UI.Menu({items: [{caption: v.txtSetProofingLanguage, value: 'lang'}]})
            });
            v.btnRevLanguage.menu.on('item:click', function () {
                var src = v.btnRevLanguage.mirrorSource();
                src && !src.isDisabled() && src.trigger('click', src);
            });

            // Comments: + Previous, Next, Show Comments
            var $comments = $rp.find('#slot-comment-remove').closest('.group').attr('data-group', 'comments');
            $comments.append('<div class="elset-column">' + small('slot-elx-rev-prevcomment') + small('slot-elx-rev-nextcomment') + '</div>' + big('slot-elx-rev-showcomments'));
            v.btnRevPrevComment = v.liveBtn($rp.find('#slot-elx-rev-prevcomment'), {cls: 'btn-toolbar', caption: v.txtPrevComment, iconCls: 'toolbar__icon btn-elx-prev-comment', hint: v.tipPrevComment});
            v.btnRevNextComment = v.liveBtn($rp.find('#slot-elx-rev-nextcomment'), {cls: 'btn-toolbar', caption: v.txtNextComment, iconCls: 'toolbar__icon btn-elx-next-comment', hint: v.tipNextComment});
            v.btnRevPrevComment.on('click', function () { me.gotoComment(-1); });
            v.btnRevNextComment.on('click', function () { me.gotoComment(1); });
            var lm = function () { var c = app.getController('LeftMenu'); return c && c.leftMenu; };
            v.btnRevShowComments = v.liveMirror($rp.find('#slot-elx-rev-showcomments'), function () { var l = lm(); return l && l.btnComments; }, {
                caption: v.txtShowComments, iconCls: 'toolbar__icon btn-big-menu-comments'
            });

            // Tracking: Track Changes (For Everyone / Just Mine) and Display for Review
            var $on = $rp.find('#btn-review-on');
            if ($on.length && rv.btnTurnOn) {
                $on.after(big('slot-elx-rev-track'));
                this.hideSource($on);
                var trackMenu = new Common.UI.Menu({items: [
                    {caption: v.txtForEveryone, value: 2, checkable: true, toggleGroup: 'elx-track'},
                    {caption: v.txtJustMine, value: 0, checkable: true, toggleGroup: 'elx-track'}
                ]});
                v.btnRevTrack = v.liveMirror($rp.find('#slot-elx-rev-track'), function () { return rv.btnTurnOn; }, {
                    caption: v.txtTrackChanges, iconCls: 'toolbar__icon btn-big-ic-review', split: !rc.appConfig || !rc.appConfig.isReviewOnly,
                    menu: (rv.btnTurnOn.options.split ? trackMenu : null)
                });
                v.btnRevTrack.options.enableToggle = true;
                v.btnRevTrack.menu && v.btnRevTrack.menu.on('show:before', function (menu) {
                    var idx = Common.Utils.InternalSettings.get('de-track-changes');
                    menu.items[0].setChecked(idx == 2, true);
                    menu.items[1].setChecked(idx == 0, true);
                });
                v.btnRevTrack.menu && v.btnRevTrack.menu.on('item:click', function (menu, item) { v.btnRevTrack.clickSourceItem(item.value); });
            }
            $rp.find('#btn-review-on, #slot-elx-rev-track').first().closest('.group').attr('data-group', 'tracking');

            // Changes: Accept, Reject, then Previous and Next
            var $prev = $rp.find('#btn-change-prev'), $next = $rp.find('#btn-change-next');
            if ($prev.length && rv.btnPrev) {
                var $chg = $prev.closest('.group');
                this.hideSource($prev);
                this.hideSource($next);
                $chg.before(sep('elx-changes-sep'));
                $chg.attr('data-group', 'changes').append('<div class="elset-column">' + small('slot-elx-rev-prev') + small('slot-elx-rev-next') + '</div>');
                v.btnRevPrev = v.liveMirror($rp.find('#slot-elx-rev-prev'), function () { return rv.btnPrev; }, {cls: 'btn-toolbar', caption: v.txtPrevChange});
                v.btnRevNext = v.liveMirror($rp.find('#slot-elx-rev-next'), function () { return rv.btnNext; }, {cls: 'btn-toolbar', caption: v.txtNextChange});
            }

            // Collaboration, Chat, History and Mail Merge after Protect (Word has none of them)
            var $tail = $();
            _.each(['#slot-btn-sharing', '#slot-btn-chat', '#slot-btn-history', '#slot-btn-mailrecepients'], function (id) {
                var $g = $rp.find(id).closest('.group');
                if ($g.length) $tail = $tail.add($g.prev('.separator')).add($g);
            });
            $tail.length && $rp.append($tail.detach());
            // tidy separators: never two in a row, none first or last
            this.tidySeparators($rp);
            tb.tagPanelGroups && tb.tagPanelGroups($rp);
            tb.clearActiveData && tb.clearActiveData('review');
        },

        tidySeparators: function ($panel) {
            var prevSep = true;
            $panel.children().each(function () {
                var $c = $(this);
                if ($c.hasClass('more-box')) return;
                if ($c.hasClass('separator')) {
                    if (prevSep) $c.remove(); else prevSep = true;
                } else if ($c.hasClass('group')) prevSep = false;
            });
            var $last = $panel.children(':not(.more-box)').last();
            $last.hasClass('separator') && $last.remove();
        },

        renameReviewControls: function () {
            var v = this.view, rc = this.getApplication().getController('Common.Controllers.ReviewChanges'),
                rv = rc && rc.getView('Common.Views.ReviewChanges');
            if (!rv || !rv.btnReviewView || !rv.btnReviewView.menu || !_.isObject(rv.btnReviewView.menu)) return;
            var names = {markup: v.txtAllMarkup, simple: v.txtSimpleMarkup, final: v.txtNoMarkup, original: v.txtOriginal};
            _.each(rv.btnReviewView.menu.items, function (it) { names[it.value] && it.setCaption(names[it.value]); });
            var sync = function () {
                var it = _.find(rv.btnReviewView.menu.items, function (i) { return i.checked; });
                it && names[it.value] && rv.btnReviewView.setCaption(names[it.value]);
            };
            rv.btnReviewView.menu.on('item:click', function () { setTimeout(sync, 0); });
            rv.btnReviewView.menu.on('show:after', sync);
            sync();
        },

        gotoComment: function (dir) {
            var api = this.api, cc = this.getApplication().getController('Common.Controllers.Comments');
            if (!api || !cc || !cc.collection) return;
            var list = cc.collection.filter(function (m) { return !m.get('hide') && !m.get('removed'); });
            if (!list.length) return;
            list = _.sortBy(list, function (m) { return m.get('position') || 0; });
            var cur = this._lastComment ? _.findIndex(list, function (m) { return m.get('uid') === this._lastComment; }, this) : -1;
            var next = cur < 0 ? (dir > 0 ? 0 : list.length - 1) : (cur + dir + list.length) % list.length;
            var id = list[next].get('uid');
            this._lastComment = id;
            api.asc_selectComment(id);
            api.asc_showComment(id);
        },

        showWordCount: function () {
            var me = this, v = this.view, api = this.api;
            if (!api) return;
            var rows = [['PageCount', v.txtPages], ['WordsCount', v.txtWords], ['SymbolsCount', v.txtCharsNoSpaces],
                        ['SymbolsWSCount', v.txtCharsSpaces], ['ParagraphCount', v.txtParagraphs]];
            var html = '<div class="elx-word-count"><div class="elx-wc-title">' + v.txtStatistics + '</div><table>' +
                _.map(rows, function (r) { return '<tr><td>' + r[1] + '</td><td class="elx-wc-val" data-key="' + r[0] + '">' + v.txtCounting + '</td></tr>'; }).join('') +
                '</table></div>';
            var win = new Common.UI.Window({
                title: v.txtWordCountTitle, width: 320, cls: 'modal-dlg', buttons: ['close'], tpl: html
            });
            win.show();
            win.$window.find('.box, .body').length;
            var info = {};
            var fill = function () {
                win.$window && win.$window.find('.elx-wc-val').each(function () {
                    var k = $(this).attr('data-key');
                    info[k] !== undefined && $(this).text(info[k].toLocaleString ? info[k].toLocaleString() : info[k]);
                });
            };
            var onInfo = function (obj) {
                if (!obj) return;
                _.each(['PageCount', 'WordsCount', 'ParagraphCount', 'SymbolsCount', 'SymbolsWSCount'], function (k) {
                    var val = obj['get_' + k]();
                    val > -1 && (info[k] = val);
                });
                fill();
            };
            api.asc_registerCallback('asc_onDocInfo', onInfo);
            api.asc_registerCallback('asc_onGetDocInfoEnd', fill);
            win.on('close', function () {
                api.asc_unregisterCallback('asc_onDocInfo', onInfo);
                api.asc_unregisterCallback('asc_onGetDocInfoEnd', fill);
            });
            api.startGetDocInfo();
        },

        // ---- View ------------------------------------------------------------------------------
        arrangeView: function () {
            var me = this, v = this.view, tb = this.toolbar,
                vc = this.getApplication().getController('ViewTab'),
                vv = vc && vc.getView('ViewTab'),
                $vp = tb.$el.find('.panel[data-tab=view]');
            if (!$vp.length || !vv) return;
            var sep = function () { return $('<div class="separator long"></div>'); },
                big = function (id) { return '<span class="btn-slot text x-huge" id="' + id + '"></span>'; };

            // Views: Read Mode, Print Layout
            var $views = $('<div class="group" data-group="views">' + big('slot-elx-view-read') + big('slot-elx-view-print') + '</div>');
            $vp.prepend($views, sep());
            var header = function () { var c = me.getApplication().getController('Viewport'); return c && c.getView('Common.Views.Header'); },
                docMode = function () { var h = header(); return h && h.btnDocMode; };
            var noViewLock = [Common.enumLock.lostConnect, Common.enumLock.disableOnStart];
            v.btnViewRead = v.liveBtn($vp.find('#slot-elx-view-read'), {caption: v.txtReadMode, iconCls: 'toolbar__icon btn-elx-read-mode', hint: v.tipReadMode, lock: noViewLock, enableToggle: true, allowDepress: false, toggleGroup: 'elx-views'});
            v.btnViewPrint = v.liveBtn($vp.find('#slot-elx-view-print'), {caption: v.txtPrintLayout, iconCls: 'toolbar__icon btn-elx-print-layout', hint: v.tipPrintLayout, lock: noViewLock, enableToggle: true, allowDepress: false, toggleGroup: 'elx-views'});
            v.btnViewPrint.toggle(true, true);
            // the same switch as the Editing/Viewing menu of the title bar
            var setMode = function (value) { Common.NotificationCenter.trigger('doc:mode-apply', value, true); };
            v.btnViewRead.on('click', function () { setMode('view'); });
            v.btnViewPrint.on('click', function () { setMode('edit'); });
            if (!docMode() || !docMode().menu) $views.add($views.next('.separator')).addClass('elx-hidden-group');
            Common.NotificationCenter.on('doc:mode-changed', function (mode) {
                v.btnViewRead.toggle(mode === 'view', true);
                v.btnViewPrint.toggle(mode !== 'view', true);
            });

            // Show: Ruler, Navigation Pane (check boxes, like Word)
            var $rul = $vp.find('#slot-chk-rulers'), $rgrp = $rul.closest('.group');
            $rgrp.attr('data-group', 'show').find('.elset').last().append('<span class="btn-slot text" id="slot-elx-view-navpane"></span>');
            v.chkNavPane = new Common.UI.CheckBox({labelText: v.txtNavigationPane, lock: [Common.enumLock.lostConnect, Common.enumLock.disableOnStart], dataHint: '1', dataHintDirection: 'left', dataHintOffset: 'small'});
            v.chkNavPane.render($vp.find('#slot-elx-view-navpane'));
            v.lockedControls.push(v.chkNavPane);
            var $nav = $vp.find('#slot-btn-navigation').closest('.group');
            if (vv.btnNavigation) {
                $nav.addClass('elx-hidden-group').next('.separator').addClass('elx-hidden-group');
                v.chkNavPane.setValue(!!vv.btnNavigation.pressed, true);
                v.chkNavPane.on('change', function (field, value) {
                    var state = field.getValue() === 'checked';
                    if (!!vv.btnNavigation.pressed !== state) { vv.btnNavigation.toggle(state); vv.btnNavigation.trigger('click', vv.btnNavigation); }
                });
                var origToggle = vv.btnNavigation.toggle;
                vv.btnNavigation.toggle = function () { origToggle.apply(this, arguments); v.chkNavPane.setValue(!!this.pressed, true); };
            } else v.chkNavPane.setVisible(false);
            // the toolbar/status bar/panels check boxes are the interface, not Word's Show group
            $vp.find('#slot-chk-toolbar').closest('.group').attr('data-group', 'window');
            // Show right after Views
            var $show = $rgrp.add($rgrp.prev('.separator'));
            $views.next('.separator').after($show.detach());
            $rgrp.after(sep());

            // Zoom: zoom box, 100%, then One Page, Multiple Pages, Page Width
            var $zoom = $vp.find('.slot-field-zoom').first().closest('.group').attr('data-group', 'zoom'),
                $ftp = $vp.find('.slot-btn-ftp').first().closest('.group'),
                $z100 = $vp.find('#slot-btn-zoom-100'), $mp = $vp.find('#slot-btn-multiple-pages'),
                $ftw = $vp.find('.slot-btn-ftw').first();
            if ($zoom.length && $ftp.length) {
                $zoom.after('<div class="group" data-group="zoom">' + big('slot-elx-view-zoom100') + '</div>');
                var $col = $ftp.addClass('elx-rows3');
                $col.append($('<div class="elset"></div>').append($mp));
                $col.append($('<div class="elset"></div>').append($ftw));
                $z100.closest('.group').addClass('elx-hidden-group');
                this.hideSource($z100);
                v.btnViewZoom100 = v.liveMirror($vp.find('#slot-elx-view-zoom100'), function () { return vv.btnZoom100; }, {
                    caption: v.txtZoom100, iconCls: 'toolbar__icon btn-elx-zoom-100'
                });
            }
            this.tidySeparators($vp);
            tb.tagPanelGroups && tb.tagPanelGroups($vp);
            tb.clearActiveData && tb.clearActiveData('view');
        },

        renameReferencesControls: function () {
            var c = this.getApplication().getController('Links'), lv = c && c.getView('Links');
            lv && lv.btnCaption && lv.btnCaption.setCaption(this.view.txtInsertCaption);
        },

        renameViewControls: function () {
            var v = this.view, vc = this.getApplication().getController('ViewTab'), vv = vc && vc.getView('ViewTab');
            if (!vv) return;
            vv.chRulers && vv.chRulers.setCaption && vv.chRulers.setCaption(v.txtRuler);
            _.each(vv.btnsFitToPage || [], function (b) { b.setCaption(v.txtOnePage); });
            _.each(vv.btnsFitToWidth || [], function (b) { b.setCaption(v.txtPageWidth); });
            vv.btnMultiplePages && vv.btnMultiplePages.setCaption(v.txtMultiplePages);
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

/*
 * SPDX-FileCopyrightText: 2026 ElixirOffice contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 *
 * PE.Controllers.SlideRibbon: PowerPoint-like ribbon of the presentation
 * editor. Tabs in PowerPoint's order (File, Home, Insert, Draw, Design,
 * Transitions, Animations, Slide Show, Review, View), PowerPoint's groups with
 * labels (ElixirRibbonLabels), Office-like collapse (Common.UI.RibbonCollapse).
 *
 * Toolbar 'render:before': the template (toolbar.$layout) is rearranged and
 * the stock controls get their PowerPoint captions before they are injected.
 * Stock controls that PowerPoint shows elsewhere stay in the layout, hidden
 * (class elx-source), so that their controllers and locks keep working.
 * Toolbar.template and the stock views are not modified.
 */
define([
    'core',
    'presentationeditor/main/app/elixir-slide/RibbonView',
    'common/main/lib/elixir-shared/RibbonLabels',
    'common/main/lib/component/RibbonCollapse'
], function () {
    'use strict';

    var TAB_ORDER = ['file', 'home', 'ins', 'draw', 'design', 'transit', 'animate', 'slideshow', 'review', 'view'];
    // PowerPoint's KeyTips for the tabs (Alt, then the letter)
    var TAB_KEYTIPS = { file: 'F', home: 'H', ins: 'N', draw: 'JI', design: 'G', transit: 'K', animate: 'A',
                        slideshow: 'S', review: 'R', view: 'W', slideMaster: 'M', plugins: 'X' };

    PE.Controllers.SlideRibbon = Backbone.Controller.extend({
        views: ['SlideRibbon'],

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

        appOptions: function () {
            var main = this.getApplication().getController('Main');
            return (main && main.appOptions) || {};
        },

        hideSource: function ($slot) { $slot.addClass('elx-source'); },

        slot: function (id, cls) {
            return $('<span class="btn-slot ' + (cls === undefined ? 'text x-huge' : cls) + '" id="' + id + '"></span>');
        },
        group: function (key, cls) {
            return $('<div class="group' + (cls ? ' ' + cls : '') + '"' + (key ? ' data-group="' + key + '"' : '') + '></div>');
        },
        sep: function () { return $('<div class="separator long"></div>'); },

        onToolbarBeforeRender: function (tb) {
            var mode = this.appOptions();
            if (!mode.isEdit || this.view) return;
            this.toolbar = tb;
            this.view = this.createView('SlideRibbon', { toolbar: tb, mode: mode });
            this.captions(tb);
            this.layoutHome(tb.$layout);
            this.layoutInsert(tb.$layout);
            this.layoutDesign(tb.$layout);
            this.layoutTransitions(tb.$layout);
            this.layoutAnimations(tb.$layout);
        },

        // PowerPoint captions on stock controls, set before they are rendered
        captions: function (tb) {
            var v = this.view;
            var cap = function (btn, text, cls, icon) {
                if (!btn) return;
                btn.options.caption = btn.caption = text;
                if (cls) btn.options.cls = btn.cls = cls;
                if (icon) btn.options.iconCls = btn.iconCls = icon;
            };
            cap(tb.btnPaste, v.txtPaste, 'btn-toolbar x-huge icon-top', 'toolbar__icon btn-big-paste');
            cap(tb.btnAddSlide, v.txtNewSlide);
            cap(tb.btnChangeSlide, v.txtLayout);
            cap(tb.btnReplace, v.txtReplace);
            cap(tb.btnSelectAll, v.txtSelectAll);
            cap(tb.btnShapeArrange, v.txtArrange);
            cap(tb.btnShapeAlign, v.txtAlign);
        },

        // ---- Home: Clipboard | Slides | Font | Paragraph | Drawing | Editing ------------
        layoutHome: function ($layout) {
            var me = this,
                $static = $layout.find('.panel.static'),
                $home = $layout.find('.panel[data-tab=home]');
            if (!$home.length) return;
            var $first = $home.children().first();

            // Clipboard: Paste (large), Cut, Copy, Format Painter
            var $clip = this.group('clipboard').append(
                $static.find('#slot-btn-paste').attr('class', 'btn-slot text x-huge'));
            var $clipSmall = this.group('', 'small').append(
                $('<div class="elset"></div>').append($static.find('#slot-btn-cut'), $static.find('#slot-btn-copy')),
                $('<div class="elset"></div>').append($static.find('#slot-btn-copystyle')));
            $first.before($clip, $clipSmall, this.sep());

            // Slides: New Slide (large, layouts gallery), Layout, Reset
            var $slides = this.group('slides').append($static.find('#slot-addslide'));
            var $slidesSmall = this.group('', 'small').append(
                $('<div class="elset"></div>').append($static.find('#slot-changeslide').removeClass('split').addClass('text')),
                $('<div class="elset"></div>').append(this.slot('slot-elx-reset-slide', 'text')));
            $first.before($slides, $slidesSmall, this.sep());

            // what is left in the static panel: Undo/Redo and Start Slide Show are in the
            // title bar and in the Slide Show tab, like PowerPoint; kept for their handlers
            $static.find('#slot-btn-undo, #slot-btn-redo, #slot-preview').each(function () { me.hideSource($(this)); });
            $static.find('.separator').addClass('elx-hidden-group');
            $static.addClass('elx-static-slide');

            // Font and Paragraph
            $first.attr('data-group', 'font');
            $first.attr('data-launcher', 'font');
            var $para = $home.find('#slot-btn-markers').closest('.group');
            $para.attr({'data-group': 'paragraph', 'data-launcher': 'paragraph'});

            // Drawing: Shapes gallery, Arrange, Align, Merge; Text Box and Picture live in Insert
            var $ins = $home.find('.slot-instext').closest('.group');
            $ins.addClass('elx-hidden-group');
            $ins.next('.separator').addClass('elx-hidden-group');
            $home.find('.slot-combo-insertshape').attr('data-group', 'drawing').attr('data-launcher', 'drawing');
            // Arrange / Align with their captions, Merge Shapes next to Align
            var $arr = $home.find('#slot-btn-arrange-shape').addClass('text');
            $home.find('#slot-btn-align-shape').addClass('text').after($home.find('#slot-btn-shapes-merge'));
            void $arr;

            // Editing: Find, Replace, Select
            var $edit = $home.find('#slot-btn-replace').closest('.group');
            $edit.attr('data-group', 'editing');
            $edit.children('.elset').first().prepend(this.slot('slot-elx-find', 'text'));
            $edit.find('#slot-btn-replace, #slot-btn-select-all').addClass('text');
            $edit.find('#slot-btn-replace').appendTo($('<div class="elset"></div>').insertAfter($edit.children('.elset').first()));
            // three rows: Find / Replace / Select
            $edit.addClass('elx-three-rows');
            $static.children('.group').each(function () {
                var $g = $(this);
                if ($g.find('.btn-slot').length && !$g.find('.btn-slot:not(.elx-source)').length) $g.addClass('elx-hidden-group');
            });
        },

        // ---- Insert: Slides | Tables | Images | Illustrations | Links | Comments | Text |
        //      Symbols | Media --------------------------------------------------------------
        layoutInsert: function ($layout) {
            var $ins = $layout.find('.panel[data-tab=ins]');
            if (!$ins.length) return;
            var $table = $ins.find('#slot-btn-inserttable').closest('.group'),
                $img = $ins.find('.slot-insertimg').closest('.group'),
                $shapes = $ins.find('.slot-combo-insertshape'),
                $chart = $ins.find('#slot-btn-insertchart').closest('.group'),
                $links = $ins.find('#slot-btn-insertlink').closest('.group'),
                $text = $ins.find('#slot-btn-instextart').closest('.group'),
                $sym = $ins.find('#slot-btn-insertequation').closest('.group'),
                $media = $ins.find('#slot-btn-insaudio').closest('.group');

            // Slides: New Slide (a second button with the layouts gallery)
            $table.before(this.group('slides').append(this.slot('slot-elx-ins-newslide')), this.sep());
            $table.attr('data-group', 'tables');
            $img.attr('data-group', 'images');
            $shapes.attr('data-group', 'illustrations');
            $shapes.prev('.separator').removeClass('invisible');
            // Comments get their own group, after Links, like PowerPoint
            var $comment = $links.find('.slot-comment');
            $links.attr('data-group', 'links');
            $links.after(this.sep(), this.group('comments').append($comment));
            $text.attr('data-group', 'text');
            // PowerPoint's order in Text: Text Box, Header & Footer, WordArt, Date & Time, Slide Number
            $text.find('#slot-btn-editheader').insertAfter($text.find('.slot-instext'));
            $sym.attr('data-group', 'symbols');
            // Symbol before Equation
            $sym.find('#slot-btn-inssymbol').prependTo($sym);
            $media.attr('data-group', 'media');
            // Video before Audio
            $media.find('#slot-btn-insvideo').prependTo($media);
            void $chart;
        },

        // ---- Design: Themes | Variants | Customize -------------------------------------
        layoutDesign: function ($layout) {
            var $d = $layout.find('.panel[data-tab=design]');
            if (!$d.length) return;
            $d.find('#slot-field-styles').attr('data-group', 'themes');
            var $g = $d.find('#slot-btn-colorschemas').closest('.group');
            // Variants (colour schemes) | Customize (Slide Size, Format Background)
            var $custom = this.group('customize').append($g.find('#slot-btn-slidesize'), this.slot('slot-elx-format-bg'));
            $g.attr('data-group', 'variants').after(this.sep(), $custom);
        },

        // ---- Transitions: Preview | Transition to This Slide | Timing --------------------
        layoutTransitions: function ($layout) {
            var $t = $layout.find('.panel[data-tab=transit]');
            if (!$t.length) return;
            var $gallery = $t.find('#transit-field-effects');
            // Preview first, as a large button
            var $prev = $t.find('#transit-button-preview');
            var $pg = this.group('preview').append($prev.addClass('x-huge'));
            $gallery.before($pg, this.sep());
            $gallery.attr('data-group', 'transition').attr('data-launcher', 'transition');
            // Timing: Duration, Apply To All | On Mouse Click, After
            var $dur = $t.find('#transit-duration').closest('.group');
            $dur.attr('data-group', 'timing');
            $dur.find('.elset').last().append($t.find('#transit-button-apply'));
            // the old Preview/Apply groups are empty now
            $t.children('.group').each(function () {
                var $g = $(this);
                if (!$g.find('.btn-slot[id]').length && !$g.is($gallery)) { $g.addClass('elx-hidden-group'); $g.prev('.separator').addClass('elx-hidden-group'); }
            });
            // Advance Slide checkboxes join Timing
            var $adv = $t.find('#transit-checkbox-startonclick').closest('.group');
            $adv.prev('.separator').addClass('invisible');
        },

        // ---- Animations: Preview | Animation | Advanced Animation | Timing ----------------
        layoutAnimations: function ($layout) {
            var $a = $layout.find('.panel[data-tab=animate]');
            if (!$a.length) return;
            var $gallery = $a.find('#animation-field-effects');
            var $prev = $a.find('#animation-button-preview');
            var $prevGroup = $prev.closest('.group');
            $gallery.before(this.group('preview').append($prev.addClass('x-huge')), this.sep());
            $prevGroup.addClass('elx-hidden-group');
            $prevGroup.prev('.separator').addClass('elx-hidden-group');
            $gallery.attr('data-group', 'animation').attr('data-launcher', 'animation');
            // Advanced Animation: Add Animation, Animation Pane, Trigger
            var $add = $a.find('#animation-button-add-effect').closest('.group');
            $add.attr('data-group', 'advAnimation');
            var $pane = $a.find('#animation-button-pane');
            var $paneGroup = $pane.closest('.group');
            $add.append($pane.addClass('x-huge'));
            $paneGroup.addClass('elx-hidden-group');
            $paneGroup.prev('.separator').addClass('elx-hidden-group');
            var $trigger = $a.find('#animation-trigger');
            $add.append($trigger);
            // Timing: Start, Duration, Delay, Repeat, Rewind, Move Earlier / Later
            $a.find('#animation-label-start').closest('.group').attr('data-group', 'timing');
            $a.find('#animation-moveearlier').closest('.group').prev('.separator').addClass('invisible');
        },

        // ---- new controls, labels, behaviour ----------------------------------------------
        onToolbarAfterRender: function (tb) {
            if (!this.view) return;
            var me = this, v = this.view, $el = tb.$el;
            var put = function (id, btn) { var $s = $el.find('#' + id); $s.length && btn.render($s); return btn; };

            v.btnResetSlide = put('slot-elx-reset-slide', v.small({ iconCls: 'toolbar__icon btn-elx-reset-slide', caption: v.txtReset, hint: v.tipReset }));
            v.btnResetSlide.on('click', function () { me.dh() && me.dh().onResetSlide(); });
            v.btnFind = put('slot-elx-find', v.small({ iconCls: 'toolbar__icon btn-menu-search', caption: v.txtFind }));
            v.btnFind.on('click', function () {
                var c = me.getApplication().getController('ElixirChrome');
                c && c.findInDocument();
            });
            v.btnInsNewSlide = put('slot-elx-ins-newslide', this.newSlideButton());
            v.btnFormatBg = put('slot-elx-format-bg', v.big({ iconCls: 'toolbar__icon btn-elx-format-bg', caption: v.txtFormatBackground, hint: v.tipFormatBackground }));
            v.btnFormatBg.on('click', function () { me.formatBackground(); });

            Array.prototype.push.apply(tb.lockControls, v.lockedControls);
            Common.Utils.lockControls(Common.enumLock.disableOnStart, true, { array: v.lockedControls });

            ElixirRibbonLabels.install(tb, {
                labels: this.labels(),
                groupTags: this.groupTags(),
                launcherTip: v.tipGroupLauncher,
                onLauncher: function (key) { me.onLauncher(key); }
            });
            Common.UI.RibbonCollapse.install(tb, { labelOf: function (key) { return tb.groupLabelText(key); } });
            this.createSlideShowTab(tb);
            this.setupRibbonPin(tb);
            var api = this.getApplication().getController('Viewport').getApi();
            this.api = api;
            // Review, View, Draw are built by their own controllers right after the toolbar
            setTimeout(function () { me.arrangeTabs(); }, 0);
        },

        dh: function () { return this.getApplication().getController('DocumentHolder'); },
        tbc: function () { return this.getApplication().getController('Toolbar'); },

        // Insert > New Slide: a second New Slide button with the same layouts gallery
        newSlideButton: function () {
            var tb = this.toolbar, v = this.view;
            var btn = v.big({ iconCls: 'toolbar__icon btn-addslide', caption: v.txtNewSlide, split: true, menu: true,
                              hint: tb.btnAddSlide && tb.btnAddSlide.options.hint });
            btn.setMenu(new Common.UI.Menu({
                items: [
                    {template: _.template('<div class="menu-layouts" style="width: 302px; margin: 0 4px;"></div>')},
                    {caption: '--'},
                    {caption: tb.txtDuplicateSlide, value: 'duplicate'}
                ]
            }));
            btn.on('click', function () { tb.fireEvent('add:slide'); });
            btn.menu.on('item:click', function (menu, item) { item.value === 'duplicate' && tb.fireEvent('duplicate:slide'); });
            // the stock gallery builder (Toolbar.updateAddSlideMenu) works on any button with this menu
            var hook = function () {
                var b = tb.binding && tb.binding.onShowBeforeAddSlide;
                if (b && !btn._elxHooked) { btn._elxHooked = true; btn.menu.on('show:before', b, btn); }
                else if (btn.mnuSlidePicker) btn.mnuSlidePicker._needRecalcSlideLayout = true;
            };
            var upd = tb.updateAddSlideMenu;
            tb.updateAddSlideMenu = function () { var r = upd.apply(this, arguments); hook(); return r; };
            return btn;
        },

        // Design > Format Background: the slide settings (fill) pane
        formatBackground: function () {
            var api = this.api, el = null;
            _.each(api ? api.getSelectedElements() : [], function (e) {
                if (e.get_ObjectType() === Asc.c_oAscTypeSelectElement.Slide) el = e;
            });
            var rm = this.getApplication().getController('RightMenu');
            rm && el && rm.onDoubleClickOnObject(el);
        },

        onLauncher: function (key) {
            var tb = this.toolbar, tbc = this.tbc();
            switch (key) {
                case 'font':
                case 'paragraph':
                    // PowerPoint's Font / Paragraph dialogs: the paragraph settings (Indents & Spacing, Font, Tabs)
                    var dh = this.dh();
                    dh && dh.onParagraphAdvanced && dh.onParagraphAdvanced();
                    break;
                case 'drawing':
                    this.rightPanel(Asc.c_oAscTypeSelectElement.Shape);
                    break;
                case 'transition':
                    tb.setTab('transit');
                    break;
                case 'animation':
                    var ac = this.getApplication().getController('Animation');
                    ac && ac.view && ac.view.btnAnimationPane && ac.view.btnAnimationPane.trigger('click', ac.view.btnAnimationPane);
                    break;
            }
            void tbc;
        },

        // ---- Slide Show: Start Slide Show | Set Up | Monitors ------------------------------
        createSlideShowTab: function (tb) {
            var me = this, v = this.view;
            var $panel = $('<section class="panel" data-tab="slideshow" role="tabpanel" aria-labelledby="slideshow">' +
                '<div class="group" data-group="startShow">' +
                    '<span class="btn-slot text x-huge" id="slot-elx-show-begin"></span>' +
                    '<span class="btn-slot text x-huge" id="slot-elx-show-current"></span>' +
                    '<span class="btn-slot text x-huge" id="slot-elx-show-presenter"></span>' +
                '</div><div class="separator long"></div>' +
                '<div class="group" data-group="setUp">' +
                    '<span class="btn-slot text x-huge" id="slot-elx-show-setup"></span>' +
                    '<span class="btn-slot text x-huge" id="slot-elx-show-hide"></span>' +
                '</div><div class="separator long"></div>' +
                '<div class="group small" data-group="monitors">' +
                    '<div class="elset"><span class="btn-slot text" id="slot-elx-show-usepresenter"></span></div>' +
                    '<div class="elset"></div>' +
                '</div></section>');
            tb.addTab({ action: 'slideshow', caption: v.txtTabSlideShow, extcls: 'canedit', layoutname: 'toolbar-slideshow', dataHintTitle: 'S' },
                $panel, tb.getLastTabIdx());
            var put = function (id, btn) { btn.render($panel.find('#' + id)); return btn; };
            v.btnShowBegin = put('slot-elx-show-begin', v.big({ iconCls: 'toolbar__icon btn-elx-show-begin', caption: v.txtFromBeginning, hint: v.tipFromBeginning }));
            v.btnShowCurrent = put('slot-elx-show-current', v.big({ iconCls: 'toolbar__icon btn-elx-show-current', caption: v.txtFromCurrent, hint: v.tipFromCurrent }));
            v.btnShowPresenter = put('slot-elx-show-presenter', v.big({ iconCls: 'toolbar__icon btn-elx-presenter', caption: v.txtPresenterView, hint: v.tipPresenterView }));
            v.btnShowSetup = put('slot-elx-show-setup', v.big({ iconCls: 'toolbar__icon btn-elx-show-setup', caption: v.txtSetUpShow, hint: v.tipSetUpShow }));
            v.btnHideSlide = put('slot-elx-show-hide', v.big({ iconCls: 'toolbar__icon btn-elx-hide-slide', caption: v.txtHideSlide, hint: v.tipHideSlide, enableToggle: true }));
            v.chUsePresenter = new Common.UI.CheckBox({ labelText: v.txtUsePresenterView, dataHint: '1', dataHintDirection: 'left', dataHintOffset: 'small' });
            v.chUsePresenter.render($panel.find('#slot-elx-show-usepresenter'));
            v.chUsePresenter.setValue(Common.localStorage.getBool('pe-elx-use-presenter', false), true);
            v.chUsePresenter.on('change', function (cb, val) { Common.localStorage.setBool('pe-elx-use-presenter', val === 'checked' || val === true); });

            var tbc = function () { return me.tbc(); };
            v.btnShowBegin.on('click', function () { me.startShow(false); });
            v.btnShowCurrent.on('click', function () { me.startShow(true); });
            v.btnShowPresenter.on('click', function () { tbc() && tbc().onPreviewItemClick(null, { value: 2 }); });
            v.btnShowSetup.on('click', function () { tbc() && tbc().onPreviewItemClick(null, { value: 3 }); });
            v.btnHideSlide.on('click', function (btn) {
                me.api && me.api.asc_HideSlides(btn.pressed);
                Common.NotificationCenter.trigger('edit:complete', tb);
            });
            var lock = [v.btnShowBegin, v.btnShowCurrent, v.btnShowPresenter, v.btnShowSetup, v.btnHideSlide];
            Array.prototype.push.apply(tb.lockControls, lock);
            Common.Utils.lockControls(Common.enumLock.disableOnStart, true, { array: lock });
        },

        // From Beginning (F5) / From Current Slide (Shift+F5); "Use Presenter View" starts the presenter
        startShow: function (current) {
            var tbc = this.tbc();
            if (!tbc || !this.api) return;
            var presenter = this.view && this.view.chUsePresenter && this.view.chUsePresenter.getValue() === 'checked';
            if (presenter) tbc.onPreview(current ? this.api.getCurrentPage() : 0, true, !!current);
            else tbc.onPreview(current ? this.api.getCurrentPage() : 0, false, !!current);
        },

        // PowerPoint's tab order and KeyTips
        arrangeTabs: function () {
            var tb = this.toolbar;
            if (!tb || !tb.$tabs) return;
            var $ul = tb.$tabs.first().parent(), v = this.view;
            var $li = function (t) { return $ul.find('> li > a[data-tab=' + t + ']').parent(); };
            var $prev = null;
            _.each(TAB_ORDER, function (t) {
                var $t = $li(t);
                if (!$t.length) return;
                $prev ? $t.insertAfter($prev) : $t.prependTo($ul);
                $prev = $t;
            });
            tb.$tabs = $ul.find('> li');
            _.each(TAB_KEYTIPS, function (k, t) { $li(t).find('> a').attr('data-hint-title', k); });
            // PowerPoint's names
            var cap = function (t, text) { var $a = $li(t).find('> a'); $a.length && text && $a.text(text).attr('data-title', text); };
            cap('review', v.txtTabReview);
            cap('animate', v.txtTabAnimations);
            tb.setVisible('slideshow', true);
            this.renameControls();
            // View: Presentation Views (Normal) | Master Views (Slide Master)
            var $sm = tb.$el.find('.panel[data-tab=view] #slot-btn-slide-master');
            if ($sm.length && !$sm.parent().is('[data-group=masterViews]')) {
                var $g = $sm.closest('.group');
                $g.after(this.sep(), this.group('masterViews').append($sm));
            }
            // the stock controller moves Paste into the Undo/Redo box of the static panel
            var $paste = tb.$el.find('#slot-btn-paste'), $clip = tb.$el.find('.panel[data-tab=home] .group[data-group=clipboard]');
            $clip.length && !$paste.parent().is($clip) && $paste.appendTo($clip);
            tb.btnPaste && tb.btnPaste.$el && tb.btnPaste.$el.find('button').attr('data-hint-direction', 'bottom');
            tb.$el.find('.panel.static > .group').each(function () {
                var $g = $(this);
                if ($g.find('.btn-slot').length && !$g.find('.btn-slot:not(.elx-source)').filter(function () { return this.style.display !== 'none'; }).length) $g.addClass('elx-hidden-group');
            });
            tb.$panels && tb.$panels.each(function () { tb.tagPanelGroups && tb.tagPanelGroups($(this)); });
            tb.processPanelVisible && tb.processPanelVisible(null, true);
        },

        // PowerPoint's words on stock controls that are already rendered (they have a caption node)
        renameControls: function () {
            var tb = this.toolbar, v = this.view, app = this.getApplication();
            var set = function (b, t) { b && b.rendered && b.caption && t && b.caption !== t && b.setCaption(t); };
            set(tb.btnInsertTextArt, v.txtWordArt);
            _.each(tb.btnsInsertImage || [], function (b) { set(b, v.txtPictures); });
            var tr = app.getController('Transitions'), an = app.getController('Animation');
            tr && tr.view && set(tr.view.btnParameters, v.txtEffectOptions);
            an && an.view && set(an.view.btnParameters, v.txtEffectOptions);
            an && an.view && set(an.view.btnAddAnimation, v.txtAddAnimation);
            an && an.view && set(an.view.btnAnimationPane, v.txtAnimationPane);
        },

        setupRibbonPin: function (tb) {
            var me = this, v = this.view;
            var $box = tb.$el.find('.box-controls').first();
            if (!$box.length) return;
            var $slot = $('<div class="elx-ribbon-pin"></div>').appendTo($box);
            v.btnRibbonPin = new Common.UI.Button({
                cls: 'btn-toolbar', iconCls: 'toolbar__icon btn-elx-ribbon-collapse',
                hint: v.tipCollapseRibbon, dataHint: '0', dataHintDirection: 'top', dataHintOffset: 'small'
            });
            v.btnRibbonPin.render($slot);
            var sync = function () {
                var folded = tb.isCompact();
                v.btnRibbonPin.setIconCls('toolbar__icon ' + (folded ? 'btn-elx-ribbon-pin' : 'btn-elx-ribbon-collapse'));
                v.btnRibbonPin.updateHint(folded ? v.tipPinRibbon : v.tipCollapseRibbon);
            };
            v.btnRibbonPin.on('click', function () {
                var tc = me.tbc();
                tc && tc.onChangeCompactView && tc.onChangeCompactView(null, !tb.isCompact());
                setTimeout(sync, 0);
            });
            tb.on('view:compact', function () { setTimeout(sync, 0); });
            sync();
        },

        onAppReady: function () {
            if (!this.view) return;
            var me = this, v = this.view, tb = this.toolbar;
            Common.Utils.lockControls(Common.enumLock.disableOnStart, false, { array: v.lockedControls.concat(
                [v.btnShowBegin, v.btnShowCurrent, v.btnShowPresenter, v.btnShowSetup, v.btnHideSlide]) });
            this.api && this.api.asc_registerCallback('asc_onFocusObject', function (objs) { me.onFocusObject(objs); });
            setTimeout(function () { me.arrangeTabs(); }, 0);
            void tb;
        },

        onFocusObject: function (objs) {
            var v = this.view, hidden = false;
            _.each(objs || [], function (o) {
                if (o.get_ObjectType() === Asc.c_oAscTypeSelectElement.Slide) {
                    var val = o.get_ObjectValue();
                    hidden = !!(val && val.get_IsHidden && val.get_IsHidden());
                }
            });
            v.btnHideSlide && v.btnHideSlide.toggle(hidden, true);
        },

        rightPanel: function (type) {
            var api = this.api, el = null;
            _.each(api ? api.getSelectedElements() : [], function (e) { if (e.get_ObjectType() === type) el = e; });
            var rm = this.getApplication().getController('RightMenu');
            rm && el && rm.onDoubleClickOnObject(el);
        },

        labels: function () {
            var v = this.view, out = {};
            _.each(PE.Views.SlideRibbon.STRINGS, function (val, k) {
                if (/^txtGroup/.test(k)) out[k.charAt(8).toLowerCase() + k.slice(9)] = v[k];
            });
            return out;
        },

        // labels of the panels built by the other controllers (Draw, Review, View, Slide Master)
        groupTags: function () {
            return {
                draw: [['#slot-btn-draw-select, #slot-btn-draw-eraser, [id^=slot-btn-draw]', 'drawingTools']],
                review: [
                    ['#slot-btn-sharing, #slot-btn-coauthmode', 'collaboration'],
                    ['.slot-comment, #slot-comment-remove', 'comments'],
                    ['#slot-btn-compare', 'compare'],
                    ['#slot-btn-chat', 'chat'],
                    ['#slot-btn-history', 'history'],
                    ['#slot-btn-add-password, #slot-btn-signature', 'protect']
                ],
                view: [
                    ['#slot-btn-normal', 'presViews'],
                    ['#slot-btn-hand-tool', 'tools'],
                    ['#slot-field-zoom', 'zoom'],
                    ['#slot-btn-interface-theme', 'appearance'],
                    ['#slot-chk-notes', 'show'],
                    ['#slot-chk-toolbar', 'window'],
                    ['#slot-btn-macros', 'macros']
                ],
                slideMaster: [
                    ['#id-toolbar-btn-add-slide-master', 'edit'],
                    ['#tlbtn-insertplaceholder', 'master'],
                    ['#slot-btn-closeslidemaster', 'close']
                ]
            };
        }
    });
});

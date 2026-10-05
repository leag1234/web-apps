/*
 * SPDX-FileCopyrightText: 2026 ElixirOffice contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 *
 * PE.Views.SlideRibbon: controls and strings of the PowerPoint-like ribbon
 * (see PE.Controllers.SlideRibbon). New controls are twins of the stock ones
 * (Common.UI.RibbonMirror) or call the existing controllers, so locks,
 * shortcuts and co-editing keep working.
 */
define([
    'core',
    'common/main/lib/component/Button',
    'common/main/lib/component/RibbonMirror'
], function () {
    'use strict';

    var STRINGS = {
        // group labels (Common.UI.RibbonGroupLabels: txtGroup<Key> for data-group="<key>")
        txtGroupUndo: 'Undo', txtGroupClipboard: 'Clipboard', txtGroupSlides: 'Slides', txtGroupFont: 'Font',
        txtGroupParagraph: 'Paragraph', txtGroupDrawing: 'Drawing', txtGroupEditing: 'Editing',
        txtGroupTables: 'Tables', txtGroupImages: 'Images', txtGroupIllustrations: 'Illustrations',
        txtGroupLinks: 'Links', txtGroupComments: 'Comments', txtGroupText: 'Text', txtGroupSymbols: 'Symbols',
        txtGroupMedia: 'Media', txtGroupThemes: 'Themes', txtGroupVariants: 'Variants', txtGroupCustomize: 'Customize',
        txtGroupPreview: 'Preview', txtGroupTransition: 'Transition to This Slide', txtGroupTiming: 'Timing',
        txtGroupAnimation: 'Animation', txtGroupAdvAnimation: 'Advanced Animation',
        txtGroupStartShow: 'Start Slide Show', txtGroupSetUp: 'Set Up', txtGroupMonitors: 'Monitors',
        txtGroupProofing: 'Proofing', txtGroupCollaboration: 'Collaboration', txtGroupProtect: 'Protect',
        txtGroupHistory: 'History', txtGroupChat: 'Chat', txtGroupCompare: 'Compare',
        txtGroupPresViews: 'Presentation Views', txtGroupMasterViews: 'Master Views', txtGroupShow: 'Show',
        txtGroupZoom: 'Zoom', txtGroupAppearance: 'Appearance', txtGroupWindow: 'Window', txtGroupMacros: 'Macros',
        txtGroupTools: 'Tools', txtGroupDrawingTools: 'Drawing Tools', txtGroupEdit: 'Edit Master', txtGroupMaster: 'Master Layout',
        txtGroupClose: 'Close',
        // tabs
        txtTabSlideShow: 'Slide Show', txtTabReview: 'Review', txtTabAnimations: 'Animations',
        // captions
        txtPaste: 'Paste', txtNewSlide: 'New Slide', txtLayout: 'Layout', txtReset: 'Reset',
        tipReset: 'Reset the position, size and formatting of the slide placeholders to their default settings.',
        txtFind: 'Find', txtReplace: 'Replace', txtSelect: 'Select', txtSelectAll: 'Select All',
        txtArrange: 'Arrange', txtAlign: 'Align', txtMerge: 'Merge Shapes', txtShapes: 'Shapes',
        txtPictures: 'Pictures', txtTextBox: 'Text Box', txtWordArt: 'WordArt',
        txtFormatBackground: 'Format Background', tipFormatBackground: 'Show the Format Background pane.',
        txtFromBeginning: 'From Beginning', tipFromBeginning: 'Start the slide show from the first slide (F5).',
        txtFromCurrent: 'From Current Slide', tipFromCurrent: 'Start the slide show from this slide (Shift+F5).',
        txtPresenterView: 'Presenter View', tipPresenterView: 'Start the slide show with Presenter View: next slide, notes and timer.',
        txtSetUpShow: 'Set Up Slide Show', tipSetUpShow: 'Set up advanced options for the slide show, such as looping.',
        txtHideSlide: 'Hide Slide', tipHideSlide: 'Hide the current slide from the slide show.',
        txtUsePresenterView: 'Use Presenter View',
        txtPreview: 'Preview', txtEffectOptions: 'Effect Options', txtApplyToAll: 'Apply To All',
        txtAddAnimation: 'Add Animation', txtAnimationPane: 'Animation Pane',
        txtNormal: 'Normal', txtSlideSorter: 'Slide Sorter', txtReadingView: 'Reading View',
        tipReadingView: 'View the presentation as a slide show that fits within the window.',
        txtNotes: 'Notes', txtFitToWindow: 'Fit to Window', txtSlideMaster: 'Slide Master',
        txtSpelling: 'Spelling', txtNewComment: 'New Comment',
        tipCollapseRibbon: 'Collapse the Ribbon (Ctrl+F1)', tipPinRibbon: 'Pin the ribbon (Ctrl+F1)',
        tipGroupLauncher: 'More options: {0}'
    };

    PE.Views.SlideRibbon = Backbone.View.extend(_.extend({

        initialize: function (options) {
            this.toolbar = options.toolbar;
            this.mode = options.mode || {};
            this.lockedControls = [];
        },

        // a big ribbon button
        big: function (opts) {
            var _set = Common.enumLock;
            var b = new Common.UI.Button(_.extend({
                cls: 'btn-toolbar x-huge icon-top',
                lock: [_set.lostConnect, _set.noSlides, _set.disableOnStart],
                dataHint: '1', dataHintDirection: 'bottom', dataHintOffset: 'small'
            }, opts));
            this.lockedControls.push(b);
            return b;
        },

        // a small ribbon button with its caption
        small: function (opts) {
            return this.big(_.extend({ cls: 'btn-toolbar' }, opts));
        },

        twin: function (src, opts) {
            var t = new Common.UI.RibbonMirror(src, opts || {});
            return t;
        },

        render: function () { return this; }

    }, STRINGS, PE.Views.SlideRibbon || {}));

    PE.Views.SlideRibbon.STRINGS = STRINGS;
});

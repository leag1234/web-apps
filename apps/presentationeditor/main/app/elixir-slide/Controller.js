/*
 * SPDX-FileCopyrightText: 2026 ElixirOffice contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 *
 * PE.Controllers.ElixirChrome: PowerPoint-like window chrome of the
 * presentation editor, built on the shared modules of
 * apps/common/main/lib/elixir-shared/ (title bar with Search (Alt+Q),
 * backstage File menu, side panels without the icon rail, comments docked on
 * the right). The upstream views are extended, not rewritten; the only hooks
 * are the controller declarations in app.js / app_dev.js and an @import in
 * app.less.
 */
define([
    'core',
    'common/main/lib/elixir-shared/Shared',
    'common/main/lib/elixir-shared/Icons',
    'common/main/lib/elixir-shared/CommandIndex',
    'common/main/lib/elixir-shared/Backstage',
    'common/main/lib/elixir-shared/TitleBar',
    'common/main/lib/elixir-shared/SidePanels'
], function () {
    'use strict';

    var formats = function () {
        var T = Asc.c_oAscFileType;
        return [
            { type: T.PPTX, ext: '.pptx', icon: 'slides', name: 'txtFmtPptx', desc: 'txtFmtPptxDesc', group: 'doc' },
            { type: T.PPTM, ext: '.pptm', icon: 'code', name: 'txtFmtPptm', desc: 'txtFmtPptmDesc', group: 'doc', onlyFor: 'pptm' },
            { type: T.POTX, ext: '.potx', icon: 'template', name: 'txtFmtPotx', desc: 'txtFmtPotxDesc', group: 'doc' },
            { type: T.PPSX, ext: '.ppsx', icon: 'slideshow', name: 'txtFmtPpsx', desc: 'txtFmtPpsxDesc', group: 'doc' },
            { type: T.ODP, ext: '.odp', icon: 'slides', name: 'txtFmtOdp', desc: 'txtFmtOdpDesc', group: 'doc' },
            { type: T.OTP, ext: '.otp', icon: 'template', name: 'txtFmtOtp', desc: 'txtFmtOtpDesc', group: 'doc' },
            { type: T.PDF, ext: '.pdf', icon: 'pdf', name: 'txtFmtPdf', desc: 'txtFmtPdfDesc', group: 'pdf' },
            { type: T.PDFA, ext: '.pdf', icon: 'pdf', name: 'txtFmtPdfa', desc: 'txtFmtPdfaDesc', group: 'pdf' },
            { type: T.PNG, ext: '.zip', icon: 'image', name: 'txtFmtPng', desc: 'txtFmtImgSlideDesc', group: 'image' },
            { type: T.JPG, ext: '.zip', icon: 'image', name: 'txtFmtJpg', desc: 'txtFmtImgSlideDesc', group: 'image' }
        ];
    };

    PE.Controllers.ElixirChrome = Backbone.Controller.extend({

        initialize: function () {
            var me = this;
            ElixirShared.init(PE, 'pe', {
                formats: formats,
                defaultFormat: Asc.c_oAscFileType.PPTX,
                defaultExt: 'pptx',
                blankKey: 'txtBlankPresentation',
                docTypesKey: 'txtPresentationTypes',
                pdfBullet1Key: 'txtPdfSlideBullet1',
                protectController: null,
                suggestions: function (tb) {
                    return [tb.btnAddSlide, tb.btnInsertTable, tb.btnsInsertImage && tb.btnsInsertImage[0],
                            tb.btnsInsertText && tb.btnsInsertText[0], tb.btnInsertChart, tb.btnSlideSize];
                }
            });
            this.addListeners({
                'FileMenu': {
                    'render:after': function () { me.backstage && me.backstage.onRender(); }
                },
                'LeftMenu': {
                    'comments:show': function () { me.titleBar && me.titleBar.setCommentsPressed(true); },
                    'comments:hide': function () { me.titleBar && me.titleBar.setCommentsPressed(false); }
                }
            });
            Common.NotificationCenter.on('app:ready', function (mode) {
                _.defer(function () { me.onAppReady(mode); });
            });
        },

        onLaunch: function () {
            var lm = this.getApplication().getController('LeftMenu');
            this.fileMenu = lm && lm.leftMenu && lm.leftMenu.menuFile;
            if (this.fileMenu)
                this.backstage = new Common.Views.ElixirBackstage({ controller: this, menu: this.fileMenu });
            this.titleBar = new Common.Views.ElixirTitleBar({ controller: this });
            ElixirSidePanels.prepare();
        },

        onAppReady: function (mode) {
            this.mode = mode;
            this.titleBar.setup(mode);
            ElixirSidePanels.setup();
        },

        leftMenu: function () {
            var lm = this.getApplication().getController('LeftMenu');
            return lm && lm.leftMenu;
        },

        canShare: function () {
            if (this.backstage) {
                this.backstage.mode = (this.fileMenu && this.fileMenu.mode) || this.mode || {};
                return this.backstage.canShare();
            }
            return false;
        },

        openBackstage: function (panel) {
            var lm = this.leftMenu();
            if (!lm) return;
            lm.showMenu('file');
            panel && this.fileMenu && this.fileMenu.selectMenu(panel);
        },

        toggleComments: function () {
            var lm = this.leftMenu();
            if (!lm || !lm.btnComments) return;
            if (lm.btnComments.pressed) {
                lm.btnComments.toggle(false);
                lm.onBtnMenuClick(lm.btnComments);
                Common.NotificationCenter.trigger('layout:changed', 'leftmenu');
                Common.NotificationCenter.trigger('edit:complete');
            } else {
                lm.showMenu('comments');
            }
        },

        // PowerPoint: Find (Ctrl+F) and Replace (Ctrl+H) open the search panel
        findInDocument: function (text, replace) {
            var lmc = this.getApplication().getController('LeftMenu');
            if (!lmc) return;
            lmc.onShowHideSearch(true, text);
            if (replace) {
                var panel = lmc.leftMenu && lmc.leftMenu.panelSearch;
                setTimeout(function () { panel && panel.focus && panel.focus('replace'); }, 100);
            }
        }
    });
});

/*
 * SPDX-FileCopyrightText: 2026 ElixirOffice contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 *
 * PE.Controllers.ElixirChrome: PowerPoint-like window chrome of the
 * presentation editor, built on the shared modules of
 * apps/common/main/lib/elixir-chrome/ (title bar with Search (Alt+Q),
 * backstage File menu, side panels without the icon rail, comments docked on
 * the right). The upstream views are extended, not rewritten; the only hooks
 * are the controller declarations in app.js / app_dev.js and an @import in
 * app.less.
 */
define([
    'core',
    'common/main/lib/elixir-chrome/Icons',
    'common/main/lib/elixir-chrome/CommandIndex',
    'common/main/lib/elixir-chrome/SidePanels',
    'presentationeditor/main/app/elixir-slide/Backstage',
    'presentationeditor/main/app/elixir-slide/TitleBar'
], function () {
    'use strict';

    PE.Controllers.ElixirChrome = Backbone.Controller.extend({

        initialize: function () {
            var me = this;
            this.addListeners({
                'FileMenu': {
                    'render:after': function () { me.backstage && me.backstage.onRender(); }
                },
                'LeftMenu': {
                    'comments:show': function () {
                        me.titleBar && me.titleBar.setCommentsPressed(true);
                        // the comments pane is docked on the right: the slide thumbnails stay, like PowerPoint
                        setTimeout(function () { me.keepThumbnails(); }, 0);
                    },
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
                this.backstage = new PE.Views.ElixirBackstage({ controller: this, menu: this.fileMenu });
            this.titleBar = new PE.Views.ElixirTitleBar({ controller: this });
            ElixirSidePanels.prepare();
        },

        onAppReady: function (mode) {
            this.mode = mode;
            this.titleBar.setup(mode);
            ElixirSidePanels.setup();
        },

        keepThumbnails: function () {
            var main = this.getApplication().getController('Main'), api = main && main.api;
            api && api.ShowThumbnails && api.ShowThumbnails(true);
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

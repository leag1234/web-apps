/*
 * SPDX-FileCopyrightText: 2026 ElixirOffice contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 *
 * ElixirOffice window chrome controller (spreadsheet editor).
 *
 * Wires the Excel-like backstage and title bar (shared with the document
 * editor, common/main/lib/elixir-chrome) into the spreadsheet editor without
 * changing the upstream views. The only hooks are the controller and module
 * declarations in app.js / app_dev.js.
 */
define([
    'core',
    'common/main/lib/elixir-chrome/Icons',
    'common/main/lib/elixir-chrome/CommandIndex',
    'spreadsheeteditor/main/app/elixir-chrome/Backstage',
    'spreadsheeteditor/main/app/elixir-chrome/TitleBar'
], function () {
    'use strict';

    SSE.Controllers.ElixirChrome = Backbone.Controller.extend({

        initialize: function () {
            var me = this;
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
                this.backstage = new SSE.Views.ElixirBackstage({ controller: this, menu: this.fileMenu });
            this.titleBar = new SSE.Views.ElixirTitleBar({ controller: this });
        },

        onAppReady: function (mode) {
            this.mode = mode;
            this.titleBar.setup(mode);
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

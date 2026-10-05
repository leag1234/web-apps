/*
 * ElixirOffice window chrome controller (document editor).
 *
 * Wires the Word-like backstage, title bar, navigation pane and theme into the
 * editor without changing the upstream views: it listens to their events and
 * extends their instances. The only hooks are the controller and module
 * declarations in app.js / app_dev.js.
 *
 * Licensed under the GNU AGPL v3 like the rest of web-apps.
 */
define([
    'core',
    'common/main/lib/elixir-chrome/Icons',
    'common/main/lib/elixir-chrome/CommandIndex',
    'documenteditor/main/app/elixir-chrome/Backstage',
    'documenteditor/main/app/elixir-chrome/TitleBar',
    'documenteditor/main/app/elixir-chrome/NavPane'
], function () {
    'use strict';

    DE.Controllers.ElixirChrome = Backbone.Controller.extend({

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
                this.backstage = new DE.Views.ElixirBackstage({ controller: this, menu: this.fileMenu });
            this.titleBar = new DE.Views.ElixirTitleBar({ controller: this });
            this.navPane = new DE.Views.ElixirNavPane({ controller: this });
            this.navPane.prepare();
        },

        onAppReady: function (mode) {
            this.mode = mode;
            this.titleBar.setup(mode);
            this.navPane.setup(mode);
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
            this.navPane && this.navPane.ready ? this.navPane.openTab('results', text) : lmc.onShowHideSearch(true, text);
            if (replace) {
                var panel = lmc.leftMenu && lmc.leftMenu.panelSearch;
                setTimeout(function () { panel && panel.focus && panel.focus('replace'); }, 100);
            }
        }
    });
});

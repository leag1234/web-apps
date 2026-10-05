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
    'documenteditor/main/app/elixir-chrome/Icons',
    'documenteditor/main/app/elixir-chrome/Backstage'
], function () {
    'use strict';

    DE.Controllers.ElixirChrome = Backbone.Controller.extend({

        initialize: function () {
            var me = this;
            this.addListeners({
                'FileMenu': {
                    'render:after': function () { me.backstage && me.backstage.onRender(); }
                }
            });
        },

        onLaunch: function () {
            var lm = this.getApplication().getController('LeftMenu');
            this.fileMenu = lm && lm.leftMenu && lm.leftMenu.menuFile;
            if (this.fileMenu)
                this.backstage = new DE.Views.ElixirBackstage({ controller: this, menu: this.fileMenu });
        }
    });
});

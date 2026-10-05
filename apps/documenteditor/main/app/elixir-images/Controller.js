/*
 * ElixirOffice pictures controller (document editor): Word's Layout Options button and
 * object anchor (LayoutOptions.js). The only hooks are the controller and module
 * declarations in app.js / app_dev.js and an @import in app.less.
 *
 * Licensed under the GNU AGPL v3 like the rest of web-apps.
 */
define([
    'core',
    'documenteditor/main/app/elixir-images/LayoutOptions'
], function () {
    'use strict';

    DE.Controllers.ElixirImages = Backbone.Controller.extend({

        initialize: function () {
            var me = this;
            Common.NotificationCenter.on('app:ready', function (mode) {
                _.defer(function () { me.onAppReady(mode); });
            });
        },

        onLaunch: function () {
            this.layoutOptions = new DE.Views.ElixirLayoutOptions({ controller: this });
        },

        onAppReady: function (mode) {
            var main = DE.getController('Main');
            if (!mode || !mode.isEdit || !main || !main.api) return;
            this.layoutOptions.setup(main.api, mode);
        }
    });
});

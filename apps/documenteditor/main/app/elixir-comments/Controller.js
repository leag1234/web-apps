/*
 * SPDX-FileCopyrightText: 2026 ElixirOffice contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 *
 * ElixirOffice comments controller (document editor): Word-style comment balloons in
 * the right margin (Margin.js). The only hooks are the controller and module
 * declarations in app.js / app_dev.js and an @import in app.less.
 */
define([
    'core',
    'documenteditor/main/app/elixir-comments/Margin'
], function () {
    'use strict';

    DE.Controllers.ElixirComments = Backbone.Controller.extend({

        initialize: function () {
            var me = this;
            Common.NotificationCenter.on('app:ready', function (mode) {
                _.defer(function () { me.onAppReady(mode); });
            });
        },

        onLaunch: function () {
            this.margin = DE.Views.ElixirCommentMargin.instance = new DE.Views.ElixirCommentMargin({ controller: this });
        },

        onAppReady: function (mode) {
            var main = DE.getController('Main'), comments = this.getApplication().getController('Common.Controllers.Comments');
            if (!mode || !main || !main.api || !comments || !comments.collection) return;
            if (!mode.canViewComments && !mode.canComments) return;
            this.margin.setup(main.api, mode, comments);
        }
    });
});

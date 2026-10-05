/*
 * (c) Copyright Ascensio System SIA 2010-2024
 *
 * This program is a free software product. You can redistribute it and/or
 * modify it under the terms of the GNU Affero General Public License (AGPL)
 * version 3 as published by the Free Software Foundation. In accordance with
 * Section 7(a) of the GNU AGPL its Section 15 shall be amended to the effect
 * that Ascensio System SIA expressly excludes the warranty of non-infringement
 * of any third-party rights.
 *
 * This program is distributed WITHOUT ANY WARRANTY; without even the implied
 * warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR  PURPOSE. For
 * details, see the GNU AGPL at: http://www.gnu.org/licenses/agpl-3.0.html
 *
 * The  interactive user interfaces in modified source and object code versions
 * of the Program must display Appropriate Legal Notices, as required under
 * Section 5 of the GNU AGPL version 3.
 *
 * All the Product's GUI elements, including illustrations and icon sets, as
 * well as technical writing content are licensed under the terms of the
 * Creative Commons Attribution-ShareAlike 4.0 International. See the License
 * terms at http://creativecommons.org/licenses/by-sa/4.0/legalcode
 *
 */

/**
 *  ElixirScreenTips.js (ElixirOffice)
 *
 *  Office ScreenTips on the ribbon and the quick access toolbar, shared by the
 *  editors: the name of the command with its shortcut in bold (the existing
 *  hint), one line saying what it does, shown after a short delay and then at
 *  once while the pointer moves along the ribbon (common/main/lib/mods/tooltip.js).
 *  The editors give the list [accessor, string key] and the strings in a
 *  subclass (DE.Controllers.ScreenTips, SSE.Controllers.ExcelScreenTips).
 */

define([
    'core'
], function () {
    'use strict';

    Common.Controllers = Common.Controllers || {};

    Common.Controllers.ElixirScreenTips = Backbone.Controller.extend({
        delay: 500,   // ms before the first ScreenTip, as in Office
        tips: [],

        initialize: function() {
            Common.NotificationCenter.on('app:ready', _.bind(this.onAppReady, this));
        },

        onAppReady: function() {
            var me = this;
            this.ready = true;
            this.apply();
            // controls created later (tabs, delayed render)
            setTimeout(function() { me.apply(); }, 1500);
            $('#toolbar').on('click', '.ribtab > a', function() { _.defer(function() { me.apply(); }); });
        },

        // accessor roots ("<root>.<path>")
        roots: function() {
            var app = this.getApplication();
            return {
                tb: function() { var c = app.getController('Toolbar'); return c && c.toolbar; },
                header: function() { var c = app.getController('Viewport'); return c && c.header; },
                ribbon: function() { var c = app.getController('RibbonTabs'); return c && c.view; }
            };
        },

        resolve: function(path) {
            var roots = this.roots();
            if (path.charAt(0) === '#')
                return $(path);
            var parts = path.split('.'), root = roots[parts[0]] && roots[parts[0]]();
            for (var i = 1; root && i < parts.length; i++)
                root = root[parts[i]];
            return root && root.cmpEl ? root.cmpEl : $();
        },

        // the elements a tooltip may be bound to: the control and its buttons or input
        targets: function($el) {
            return $el.add($el.find('button, input, [data-hint]'));
        },

        apply: function() {
            if (!this.ready) return;
            var me = this;
            // Office's delay on every ribbon and title bar control
            this.targets($('#toolbar [data-hint], #toolbar button')).attr('data-screentip-delay', this.delay);
            _.each(this.tips, function(entry) {
                var text = me[entry[1]];
                var $el = me.resolve(entry[0]);
                if (!text || !$el.length) return;
                me.targets($el).attr({'data-screentip': text, 'data-screentip-delay': me.delay});
            });
        }
    });
});

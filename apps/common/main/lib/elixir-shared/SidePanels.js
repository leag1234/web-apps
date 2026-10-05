/*
 * SPDX-FileCopyrightText: 2026 ElixirOffice contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 *
 * ElixirSidePanels: Office-like side panels for any editor.
 *
 * - no icon rail on the left (body.elx-norail): the left column takes no room
 *   while none of its panels is open;
 * - the Comments pane is docked on the right of the editing area, like Office.
 *
 * Generalized from the rail and comments-dock part of
 * documenteditor/main/app/elixir-chrome/NavPane.js (which keeps its own copy
 * for the Word navigation pane). Styles: elixir-shared/chrome.less.
 *
 *   ElixirSidePanels.prepare();          // before the document loads
 *   ElixirSidePanels.setup();            // on app:ready
 */
define([
    'core'
], function () {
    'use strict';

    var SidePanels = window.ElixirSidePanels = {
        prepare: function () {
            $('body').addClass('elx-norail');
        },

        setup: function () {
            var $left = $('#left-menu'), $panel = $left.find('.left-panel');
            if (!$panel.length || this.ready) return;
            this.ready = true;
            this.$left = $left;
            this.$panel = $panel;
            this.dockComments();
            this.watchLeftWidth();
        },

        layoutChanged: function () {
            _.defer(function () { Common.NotificationCenter.trigger('layout:changed', 'leftmenu'); });
        },

        watchLeftWidth: function () {
            var me = this, el = this.$left[0];
            var check = this.checkLeft = function () {
                var w = parseInt(el.style.width) || 0;
                // closed, or open for a panel that lives elsewhere (the docked comments)
                var collapsed = w <= 40 || !me.$panel.children('.content-box').filter(function () { return this.style.display !== 'none'; }).length;
                if (collapsed !== me.$left.hasClass('elx-left-collapsed')) {
                    me.$left.toggleClass('elx-left-collapsed', collapsed);
                    me.layoutChanged();
                }
            };
            new MutationObserver(check).observe(el, { attributes: true, attributeFilter: ['style'] });
            var timer;
            new MutationObserver(function () { clearTimeout(timer); timer = setTimeout(check, 0); })
                .observe(this.$panel[0], { attributes: true, attributeFilter: ['style'], subtree: true });
            check();
        },

        dockComments: function () {
            var me = this, $c = $('#left-panel-comments');
            if (!$c.length) return;
            this.$dock = $('<div id="elx-comments-dock" class="elx-comments-dock side-panel" style="display:none;"></div>').appendTo('#editor-container');
            this.$dock.append($c);
            var sync = function () {
                var open = $c[0].style.display !== 'none';
                if (open === me.$dock.is(':visible')) return;
                me.$dock.toggle(open);
                $('#editor-container').toggleClass('elx-comments-docked', open);
                me.checkLeft && me.checkLeft();
                me.layoutChanged();
            };
            new MutationObserver(sync).observe($c[0], { attributes: true, attributeFilter: ['style'] });
            sync();
        }
    };

    return SidePanels;
});

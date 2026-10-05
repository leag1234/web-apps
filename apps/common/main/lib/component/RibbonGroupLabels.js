/*
 * SPDX-FileCopyrightText: 2026 ElixirOffice contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 *
 * Common.UI.RibbonGroupLabels: Office-like ribbon group labels and dialog
 * launchers, mixed into an editor's Toolbar view (a Common.UI.Mixtbar).
 *
 * Each run of groups between two separators is a ribbon "group" in the Office
 * sense. Its label comes from the data-group attribute of its first group
 * (set in Toolbar.template, or by the view's groupTags for the panels built by
 * other controllers) and the view's txtGroup<Key> string; it is drawn by CSS
 * under the run (::after, width from --group-label-width), so the overflow
 * logic of Mixtbar, which moves the panel children around, is not disturbed.
 * A data-launcher attribute adds the launcher arrow, which fires
 * 'group:launcher' with its value. Labels are hidden unless the theme sets
 * --toolbar-group-label-display (common/main/resources/less/elixir-ribbon.less).
 */
define([
    'common/main/lib/component/Mixtbar'
], function () {
    'use strict';

    Common.UI.RibbonGroupLabels = {
        groupLabelText: function (key) {
            var name = 'txtGroup' + key.charAt(0).toUpperCase() + key.slice(1);
            return this[name] || '';
        },

        tagPanelGroups: function ($panel) {
            var me = this,
                tab = $panel && $panel.attr('data-tab'),
                tags = tab && me.groupTags && me.groupTags[tab];
            if (!tags) return;
            _.each(tags, function (t) {
                var $g = $panel.find(t[0]).first().closest('.group');
                $g.length && !$g.attr('data-group') && $g.attr('data-group', t[1]);
            });
        },

        addTab: function (tab, panel, after) {
            Common.UI.Mixtbar.prototype.addTab.apply(this, arguments);
            panel && this.tagPanelGroups(panel);
        },

        processPanelVisible: function () {
            var me = this;
            Common.UI.Mixtbar.prototype.processPanelVisible.apply(me, arguments);
            me.updateGroupLabels();
            // Mixtbar re-runs the layout pass on a 100ms timer: follow it
            clearTimeout(me._groupLabelsTimer);
            me._groupLabelsTimer = setTimeout(function () { me.updateGroupLabels(); }, 250);
        },

        measureGroupLabel: function (text, $panel) {
            var ctx = this._labelCtx || (this._labelCtx = document.createElement('canvas').getContext('2d'));
            ctx.font = '11px ' + window.getComputedStyle($panel[0]).fontFamily;
            return Math.ceil(ctx.measureText(text).width);
        },

        updateGroupLabels: function ($panel) {
            var me = this;
            $panel = $panel || (me.$panels && me.$panels.filter('.active'));
            if (!$panel || !$panel.length || !$panel.is(':visible')) return;
            var rtl = Common.UI.isRTL(),
                segments = [], cur = [];
            $panel.children().each(function () {
                var $c = $(this);
                if ($c.hasClass('separator') && !$c.hasClass('invisible')) {
                    if ($c.css('display') !== 'none') { cur.length && segments.push(cur); cur = []; }
                } else if ($c.hasClass('group') && $c.css('display') !== 'none') {
                    cur.push($c);
                }
            });
            cur.length && segments.push(cur);
            _.each(segments, function (seg) {
                var key, $owner = seg[0];
                _.each(seg, function ($g, i) {
                    $g.removeClass('group-label-owner');
                    !key && (key = $g.attr('data-group'));
                });
                var text = key ? me.groupLabelText(key) : '';
                var launcher;
                _.each(seg, function ($g) { launcher = launcher || $g.attr('data-launcher'); });
                // a label wider than its single group widens the group, like in Word
                if (seg.length === 1 && text && !$owner.hasClass('flex')) {
                    $owner.css({'width': '', 'text-align': ''});
                    var need = me.measureGroupLabel(text, $panel) + (launcher ? 36 : 8);
                    if ($owner.outerWidth() < need) $owner.css({'width': need + 'px', 'text-align': 'center'});
                }
                var first = seg[0][0].getBoundingClientRect(),
                    last = seg[seg.length - 1][0].getBoundingClientRect(),
                    width = Math.round(rtl ? first.right - last.left : last.right - first.left);
                $owner.addClass('group-label-owner').toggleClass('has-launcher', !!(launcher && text)).attr('data-label', text);
                $owner[0].style.setProperty('--group-label-width', width + 'px');
                var $btn = $owner.children('.group-launcher');
                if (launcher && text) {
                    if (!$btn.length) {
                        $btn = $('<button type="button" class="group-launcher" data-hint="1" data-hint-direction="bottom" data-hint-offset="small"></button>');
                        $btn.on('click', function (e) {
                            me.fireEvent('group:launcher', [launcher]);
                        });
                        $owner.prepend($btn);
                    }
                    $btn.attr({'aria-label': text, 'data-launcher': launcher});
                    $btn.attr('title', me.tipGroupLauncher ? Common.Utils.String.format(me.tipGroupLauncher, text) : text);
                } else {
                    $btn.remove();
                }
            });
        }
    };

    return Common.UI.RibbonGroupLabels;
});

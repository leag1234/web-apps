/*
 * SPDX-FileCopyrightText: 2026 ElixirOffice contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 *
 * ElixirRibbonLabels: Office-like group labels and dialog launchers for any
 * editor's ribbon (Common.UI.Mixtbar instance).
 *
 * Shared copy of the mechanism of DE.Views.Toolbar.updateGroupLabels
 * (documenteditor/main/app/view/Toolbar.js), installed on a toolbar instance
 * instead of being written into the editor's Toolbar view:
 *
 *   ElixirRibbonLabels.install(toolbar, {
 *       labels: {clipboard: 'Clipboard', ...},     // key -> caption
 *       groupTags: {review: [[selector, key], ...]}, // panels built by other controllers
 *       onLauncher: function (key) {...}
 *   });
 *
 * Each run of groups between two separators is a ribbon group in the Office
 * sense. Its label comes from the data-group attribute of its first group and
 * is drawn by CSS under the run (elixir-shared/ribbon.less); the overflow logic
 * of Mixtbar is not disturbed. A data-launcher attribute adds a launcher.
 */
define([
    'core'
], function () {
    'use strict';

    var Labels = window.ElixirRibbonLabels = {
        install: function (tb, opts) {
            if (!tb || tb._elxLabels) return;
            opts = opts || {};
            tb._elxLabels = true;
            tb.elxGroupTags = opts.groupTags || {};
            tb.elxLabels = opts.labels || {};

            tb.groupLabelText = function (key) {
                if (!key) return '';
                if (tb.elxLabels[key]) return tb.elxLabels[key];
                var name = 'txtGroup' + key.charAt(0).toUpperCase() + key.slice(1);
                return tb[name] || '';
            };

            tb.tagPanelGroups = function ($panel) {
                var tab = $panel && $panel.attr('data-tab'),
                    tags = tab && tb.elxGroupTags[tab];
                if (!tags) return;
                _.each(tags, function (t) {
                    var $g = $panel.find(t[0]).first().closest('.group');
                    $g.length && !$g.attr('data-group') && $g.attr('data-group', t[1]);
                });
            };

            var addTab = tb.addTab;
            tb.addTab = function (tab, panel) {
                var r = addTab.apply(this, arguments);
                panel && tb.tagPanelGroups(panel);
                return r;
            };

            var ppv = tb.processPanelVisible;
            tb.processPanelVisible = function () {
                var r = ppv.apply(this, arguments);
                Labels.update(tb, null, opts);
                // Mixtbar re-runs the layout pass on a 100ms timer: follow it
                clearTimeout(tb._groupLabelsTimer);
                tb._groupLabelsTimer = setTimeout(function () { Labels.update(tb, null, opts); }, 250);
                return r;
            };
            tb.updateGroupLabels = function ($panel) { Labels.update(tb, $panel, opts); };
        },

        measure: function (tb, text, $panel) {
            var ctx = tb._labelCtx || (tb._labelCtx = document.createElement('canvas').getContext('2d'));
            ctx.font = '11px ' + window.getComputedStyle($panel[0]).fontFamily;
            return Math.ceil(ctx.measureText(text).width);
        },

        update: function (tb, $panel, opts) {
            $panel = $panel || (tb.$panels && tb.$panels.filter('.active'));
            if (!$panel || !$panel.length || !$panel.is(':visible')) return;
            var rtl = Common.UI.isRTL(), segments = [], cur = [];
            $panel.children().each(function () {
                var $c = $(this);
                if ($c.hasClass('separator') && !$c.hasClass('invisible')) {
                    if ($c.css('display') !== 'none') { cur.length && segments.push(cur); cur = []; }
                } else if ($c.hasClass('group') && $c.css('display') !== 'none' && !$c.hasClass('elx-hidden-group')) {
                    cur.push($c);
                }
            });
            cur.length && segments.push(cur);
            _.each(segments, function (seg) {
                var key, $owner = seg[0], launcher;
                _.each(seg, function ($g) {
                    $g.removeClass('group-label-owner');
                    !key && (key = $g.attr('data-group'));
                    launcher = launcher || $g.attr('data-launcher');
                });
                var text = key ? tb.groupLabelText(key) : '';
                // a label wider than its single group widens the group, like in Office
                if (seg.length === 1 && text && !$owner.hasClass('flex')) {
                    $owner.css({'width': '', 'text-align': ''});
                    var need = Labels.measure(tb, text, $panel) + (launcher ? 36 : 8);
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
                        $btn.on('click', function () {
                            var l = $(this).attr('data-launcher');
                            opts && opts.onLauncher ? opts.onLauncher(l) : tb.fireEvent('group:launcher', [l]);
                        });
                        $owner.prepend($btn);
                    }
                    $btn.attr({'aria-label': text, 'data-launcher': launcher});
                    $btn.attr('title', opts && opts.launcherTip ? Common.Utils.String.format(opts.launcherTip, text) : text);
                } else {
                    $btn.remove();
                }
            });
        }
    };

    return Labels;
});

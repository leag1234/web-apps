/*
 * SPDX-FileCopyrightText: 2026 ElixirOffice contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 *
 * Common.UI.RibbonCollapse: Office-like responsive ribbon.
 *
 * When the active ribbon panel is too wide for the window, its groups (the runs
 * of .group between two separators, as labelled by the toolbar) collapse into a
 * single drop-down button carrying the group name, from right to left, like
 * Office. Clicking that button opens the group's own controls in a pop-up below
 * it. Groups expand again, from left to right, when room comes back. The stock
 * overflow of Mixtbar (captions hidden, then a "More" pop-up) remains the last
 * resort when even the collapsed panel does not fit.
 *
 * The controls are moved, not copied: their events, locks and state are kept.
 *
 * Common.UI.RibbonCollapse.install(toolbar, {labelOf: function(key) {...}})
 */

define([
    'core',
    'common/main/lib/component/Button'
], function () {
    'use strict';

    var Collapse = function (toolbar, opts) {
        this.toolbar = toolbar;
        this.opts = opts || {};
        this.state = {}; // tab -> [segment]
    };

    _.extend(Collapse.prototype, {
        // runs of groups between visible separators; the more-box is not part of the panel
        segments: function ($panel) {
            var segs = [], cur = null;
            $panel.children().each(function () {
                var $c = $(this);
                if ($c.hasClass('more-box') || $c.hasClass('elx-collapsed-group')) return;
                if ($c.hasClass('separator')) {
                    if ($c.css('display') === 'none' || $c.hasClass('elx-hidden-group')) { cur && cur.nodes.push(this); return; }
                    cur && segs.push(cur);
                    cur = {nodes: [this], groups: []};
                    return;
                }
                if (!$c.hasClass('group')) return;
                if (!cur) cur = {nodes: [], groups: []};
                cur.nodes.push(this);
                $c.css('display') !== 'none' && !$c.hasClass('elx-hidden-group') && cur.groups.push(this);
            });
            cur && segs.push(cur);
            return _.filter(segs, function (s) { return s.groups.length > 0; });
        },

        key: function (seg) {
            var k;
            _.each(seg.groups, function (g) { k = k || $(g).attr('data-group'); });
            return k;
        },

        // width the panel needs; flex groups (galleries) count for their minimum width
        need: function ($panel) {
            var w = 0;
            $panel.children().each(function () {
                var $c = $(this);
                if ($c.hasClass('more-box') || $c.css('display') === 'none') return;
                if ($c.hasClass('flex')) w += (parseFloat($c.css('min-width')) || 0) + (parseFloat($c.css('padding-left')) || 0) + (parseFloat($c.css('padding-right')) || 0);
                else w += $c.outerWidth(true);
            });
            return w;
        },

        available: function ($panel) {
            var $box = $panel.closest('.box-controls'),
                stat = $box.find('.panel.static').outerWidth() || 0;
            return $box.width() - stat - 12;
        },

        collapse: function (tab, $panel, seg) {
            var me = this, key = this.key(seg),
                label = (key && this.opts.labelOf && this.opts.labelOf(key)) || '',
                icon = $(seg.groups).filter('[data-collapse-icon]').attr('data-collapse-icon') || $(seg.groups).find('.btn-slot:not(.elx-source) svg use').filter(function () { return $(this).closest('.elx-source').length === 0; }).first().attr('href'),
                width = 0;
            _.each(seg.nodes, function (n) { width += $(n).outerWidth(true); });
            var $ph = $('<div class="group elx-collapsed-group"><span class="btn-slot text x-huge"></span></div>'),
                $sep = $(seg.nodes[0]).hasClass('separator') ? $('<div class="separator long elx-collapsed-sep"></div>') : null,
                $pop = $('<div class="dropdown-menu more-container elx-group-popup" data-tab="' + tab + '"><div class="elx-group-popup-inner"></div></div>');
            var btn = new Common.UI.Button({
                cls: 'btn-toolbar x-huge icon-top',
                iconCls: 'toolbar__icon ' + (icon ? icon.replace('#', '') : 'btn-big-more'),
                caption: label || Common.Locale.get('textMoreButton', {name: 'Common.Translation', default: 'More'}),
                enableToggle: true,
                menu: false,
                dataHint: '1', dataHintDirection: 'bottom', dataHintOffset: 'small'
            });
            $(seg.nodes[0]).before($ph);
            $sep && $ph.before($sep);
            btn.render($ph.find('.btn-slot'));
            $ph.find('.inner-box-caption .caption').append('<i class="caret"></i>');
            var $inner = $pop.find('.elx-group-popup-inner');
            _.each(seg.nodes, function (n) {
                if ($(n).hasClass('separator') && n === seg.nodes[0]) { $(n).addClass('elx-moved-sep').detach(); return; }
                $inner.append(n);
            });
            this.toolbar.$el.find('.toolbar').first().append($pop);
            var entry = {seg: seg, width: width, $ph: $ph, $sep: $sep, $pop: $pop, btn: btn, firstSep: $(seg.nodes[0]).hasClass('separator') ? seg.nodes[0] : null};
            btn.on('toggle', function (b, state) { state ? me.showPopup(entry) : me.hidePopup(entry); });
            (this.state[tab] || (this.state[tab] = [])).push(entry);
        },

        expand: function (tab, entry) {
            this.hidePopup(entry);
            var $anchor = entry.$sep || entry.$ph;
            entry.firstSep && $anchor.before(entry.firstSep);
            _.each(entry.seg.nodes, function (n) { n !== entry.firstSep && $anchor.before(n); });
            entry.$sep && entry.$sep.remove();
            entry.$ph.remove();
            entry.btn.remove && entry.btn.remove();
            entry.$pop.remove();
            this.state[tab] = _.without(this.state[tab], entry);
        },

        showPopup: function (entry) {
            var me = this;
            this.hideAll(entry);
            var $bar = entry.$pop.parent(), off = Common.Utils.getOffset(entry.btn.$el), parent = Common.Utils.getOffset($bar);
            entry.$pop.css({left: Math.max(4, off.left - parent.left), right: 'auto', top: off.top - parent.top + entry.btn.$el.outerHeight() + 4, display: 'block', 'max-width': Common.Utils.innerWidth() - 8});
            var over = Common.Utils.getOffset(entry.$pop).left + entry.$pop.outerWidth() - Common.Utils.innerWidth() + 4;
            over > 0 && entry.$pop.css('left', Math.max(4, parseFloat(entry.$pop.css('left')) - over));
            this._outside = function (e) {
                if ($(e.target).closest(entry.$pop).length || $(e.target).closest(entry.btn.$el).length || $(e.target).closest('.dropdown-menu').length) return;
                entry.btn.toggle(false);
            };
            setTimeout(function () { $(document).on('mousedown.elxcollapse', me._outside); }, 0);
            // a command run from the pop-up closes it, like Office
            this._done = function () { setTimeout(function () { !entry.$pop.find('.dropdown-menu:visible').length && entry.btn.pressed && entry.btn.toggle(false); }, 0); };
            entry.$pop.on('click.elxcollapse', '.btn-toolbar:not(.dropdown-toggle), .checkbox-indeterminate, label', this._done);
        },

        hidePopup: function (entry) {
            entry.$pop.hide();
            entry.$pop.off('click.elxcollapse');
            $(document).off('mousedown.elxcollapse');
            entry.btn.pressed && entry.btn.toggle(false, true);
        },

        hideAll: function (except) {
            var me = this;
            _.each(this.state, function (list) { _.each(list, function (e) { e !== except && e.$pop.is(':visible') && me.hidePopup(e); }); });
        },

        // expand, from left to right, every collapsed group that fits again
        expandFitting: function (tab, $panel, avail) {
            var me = this, changed = false,
                order = _.sortBy((this.state[tab] || []).slice(), function (e) {
                    // groups marked data-collapse-first come back last
                    return ($(e.seg.groups).filter('[data-collapse-first]').length ? 100000 : 0) + e.$ph.index();
                });
            _.each(order, function (e) {
                var need = me.need($panel) - (e.$ph.outerWidth(true) + (e.$sep ? e.$sep.outerWidth(true) : 0)) + e.width;
                if (need > avail && !$(e.seg.groups).filter('[data-collapse-first]').length) {
                    // make room by collapsing the groups marked data-collapse-first (galleries) instead
                    var first = _.filter(me.segments($panel), function (sg) { return $(sg.groups).filter('[data-collapse-first]').length; }),
                        saving = 0;
                    _.each(first, function (sg) { _.each(sg.nodes, function (n) {
                        var $n = $(n);
                        saving += $n.hasClass('flex') ? (parseFloat($n.css('min-width')) || 0) : $n.outerWidth(true);
                    }); saving -= 72; });
                    if (first.length && need - saving <= avail) {
                        _.each(first, function (sg) { me.collapse(tab, $panel, sg); });
                        need = me.need($panel) - (e.$ph.outerWidth(true) + (e.$sep ? e.$sep.outerWidth(true) : 0)) + e.width;
                    }
                }
                if (need > avail) return;
                me.expand(tab, e);
                changed = true;
            });
            return changed;
        },

        // one pass on the active panel: collapse from the right, expand from the left
        update: function ($panel) {
            if (!$panel || !$panel.length || !$panel.is(':visible') || this.busy) return false;
            var tab = $panel.attr('data-tab'), changed = false, guard = 0;
            this.busy = true;
            try {
                var avail = this.available($panel);
                changed = this.expandFitting(tab, $panel, avail) || changed;
                // collapse while it does not fit (the right-most expanded group first)
                guard = 0;
                while (this.need($panel) > avail && guard++ < 50) {
                    var segs = this.segments($panel);
                    if (segs.length < 2) break; // the first group stays, the stock overflow takes over
                    // a group marked data-collapse-first (e.g. a wide gallery) goes before the right-most one
                    var pick = segs[segs.length - 1];
                    for (var i = segs.length - 1; i > 0; i--) {
                        if ($(segs[i].groups).filter('[data-collapse-first]').length) { pick = segs[i]; break; }
                    }
                    this.collapse(tab, $panel, pick);
                    changed = true;
                }
                // a big group may have gone: smaller ones collapsed before it can come back
                changed && this.expandFitting(tab, $panel, avail);
            } finally {
                this.busy = false;
            }
            return changed;
        }
    });

    Common.UI.RibbonCollapse = {
        install: function (toolbar, opts) {
            if (toolbar._elxCollapse) return toolbar._elxCollapse;
            var c = new Collapse(toolbar, opts), orig = toolbar.processPanelVisible, origSetTab = toolbar.setTab;
            toolbar._elxCollapse = c;
            toolbar.processPanelVisible = function (panel, reset, force) {
                if (!c.disabled) {
                    var $active = panel || (this.$panels && this.$panels.filter('.active'));
                    if ($active && c.update($active)) {
                        this.clearActiveData && this.clearActiveData($active.attr('data-tab'));
                        reset = true;
                    }
                }
                return orig.call(this, panel, reset, force);
            };
            toolbar.setTab = function () {
                c.hideAll();
                return origSetTab.apply(this, arguments);
            };
            Common.NotificationCenter.on('window:resize', function () { c.hideAll(); });
            return c;
        }
    };

    return Common.UI.RibbonCollapse;
});

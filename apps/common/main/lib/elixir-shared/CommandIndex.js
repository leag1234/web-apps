/*
 * ElixirOffice window chrome: command index for the title bar search
 * (Word's "Search (Alt+Q)" / "Tell me").
 *
 * Every Common.UI.Button is remembered through a WeakRef when it is created;
 * the index keeps the ones that live in the ribbon (#toolbar) and their menu
 * items, with the ribbon tab and group they belong to. Running a command
 * switches to its tab and clicks the control, so the result is exactly what
 * the ribbon does (dialogs, menus, galleries).
 *
 * Licensed under the GNU AGPL v3 like the rest of web-apps.
 */
define([
    'core',
    'common/main/lib/component/Button'
], function () {
    'use strict';

    var registry = [];
    var hasWeakRef = typeof window.WeakRef === 'function';

    // remember the buttons as they are created (before the ribbon is built)
    (function hookButtons() {
        var proto = Common.UI.Button && Common.UI.Button.prototype;
        if (!proto || proto._elxIndexed) return;
        proto._elxIndexed = true;
        var init = proto.initialize;
        proto.initialize = function () {
            var r = init.apply(this, arguments);
            registry.push(hasWeakRef ? new WeakRef(this) : this);
            if (registry.length > 4000) prune();
            return r;
        };
    })();

    function prune() {
        registry = registry.filter(function (ref) {
            var b = hasWeakRef ? ref.deref() : ref;
            return b && b.cmpEl && document.body.contains(b.cmpEl[0]) || (b && !b.rendered);
        });
    }

    var fold = function (s) {
        s = (s || '').toString().toLowerCase();
        return s.normalize ? s.normalize('NFD').replace(/[̀-ͯ]/g, '') : s;
    };
    var clean = function (s) {
        if (!s || typeof s !== 'string') return '';
        return s.replace(/<[^>]*>/g, ' ').replace(/\s*\((?:[^()]*(?:Ctrl|Alt|Shift|⌘|⌥|⇧|F\d)[^()]*)\)\s*$/i, '')
            .replace(/&nbsp;/g, ' ').replace(/\s+/g, ' ').trim();
    };
    var hintOf = function (o) {
        var h = o && (o.options ? o.options.hint : o.hint);
        if (_.isArray(h)) h = h[0];
        return clean(typeof h === 'string' ? h : '');
    };

    var CommandIndex = {
        build: function (toolbar) {
            var cmds = [], seen = {};
            var tabCaption = function (tab) {
                var $a = toolbar && toolbar.$el ? toolbar.$el.find('.ribtab > a[data-tab="' + tab + '"]') : $();
                return $a.length ? $.trim($a.text()) : '';
            };
            // ribbon group label (toolbar.groupLabelText, see RibbonLabels): the key sits on the first group of a run
            var groupLabel = function ($el) {
                var $g = $el.closest('.group'), key = $g.attr('data-group');
                for (var $p = $g.prev(); !key && $p.length && !$p.hasClass('separator'); $p = $p.prev())
                    key = $p.attr('data-group');
                return key && toolbar && toolbar.groupLabelText ? (toolbar.groupLabelText(key) || '') : '';
            };
            var add = function (c) {
                var key = fold(c.label) + '|' + c.tab + '|' + (c.parent || '');
                if (!c.label || seen[key]) return;
                seen[key] = true;
                c.text = fold([c.label, c.hint, c.parent, c.group, c.tabCaption].join(' '));
                cmds.push(c);
            };
            registry.forEach(function (ref) {
                var b = hasWeakRef ? ref.deref() : ref;
                if (!b || !b.cmpEl || !b.cmpEl.length) return;
                var $panel = b.cmpEl.closest('#toolbar section.panel[data-tab]');
                if (!$panel.length) return;
                var tab = $panel.attr('data-tab');
                if (tab === 'file') return;
                var caption = clean(b.options && b.options.caption || b.caption),
                    hint = hintOf(b),
                    label = hint || caption;
                if (!label) return;
                var base = { btn: b, tab: tab, tabCaption: tabCaption(tab), group: groupLabel(b.cmpEl), hint: caption !== label ? caption : '' };
                add(_.extend({ label: label }, base));
                var menu = b.menu;
                if (menu && typeof menu === 'object' && menu.items) {
                    menu.items.forEach(function (item) {
                        if (!item || item.menu || item.options && (item.options.caption === '--' || item.options.template)) return;
                        var ic = clean(item.caption || item.options && item.options.caption);
                        if (!ic || ic === '--') return;
                        add(_.extend({}, base, { label: ic, parent: label, item: item, hint: '' }));
                    });
                }
            });
            this.cmds = cmds;
            return cmds;
        },

        search: function (query, limit) {
            var words = fold(query).split(/\s+/).filter(Boolean);
            if (!words.length || !this.cmds) return [];
            var res = [];
            this.cmds.forEach(function (c) {
                var l = fold(c.label), score = 0;
                for (var i = 0; i < words.length; i++) {
                    var w = words[i], idx = c.text.indexOf(w);
                    if (idx < 0) return;
                    if (l.indexOf(w) === 0) score += 6;
                    else if (new RegExp('(^|[\\s\\-/])' + w.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')).test(l)) score += 4;
                    else if (l.indexOf(w) > 0) score += 2;
                    else score += 1;
                }
                if (c.item) score -= 0.5;
                if (isDisabled(c)) score -= 3;
                res.push({ c: c, s: score });
            });
            res.sort(function (a, b) { return b.s - a.s || a.c.label.length - b.c.label.length; });
            return res.slice(0, limit || 8).map(function (r) { return r.c; });
        },

        isDisabled: function (c) { return isDisabled(c); },

        run: function (c, toolbar) {
            if (!c || isDisabled(c)) return false;
            var b = c.btn;
            if (toolbar && c.tab && toolbar.setTab) {
                if (toolbar.isTabActive ? !toolbar.isTabActive(c.tab) : true) toolbar.setTab(c.tab);
            }
            var click = function ($el) { $el && $el[0] && $el[0].click(); };
            _.defer(function () {
                if (c.item) {
                    var item = c.item;
                    var openAndPick = function () {
                        if (item.rendered && item.$el && item.$el.is(':visible')) { click(item.$el.find('> a')); return true; }
                        return false;
                    };
                    var $toggle = b.cmpEl.find('[data-toggle=dropdown]').last();
                    if (!$toggle.length) $toggle = b.cmpEl.find('button').last();
                    click($toggle);
                    setTimeout(function () {
                        if (!openAndPick()) {
                            // menu not shown (e.g. folded ribbon): fire the item like a click would
                            item.trigger('click', item, {});
                            b.menu && !b.menu.rendered && b.menu.trigger('item:click', b.menu, item, {});
                        }
                    }, 60);
                } else {
                    var $btn = b.cmpEl.is('button') ? b.cmpEl : b.cmpEl.find('button').first();
                    if (b.menu && !b.split) $btn = b.cmpEl.find('[data-toggle=dropdown]').first().length ? b.cmpEl.find('[data-toggle=dropdown]').first() : $btn;
                    click($btn);
                }
            });
            return true;
        }
    };

    function isDisabled(c) {
        var b = c.btn;
        if (!b || (b.isDisabled && b.isDisabled())) return true;
        if (c.item && c.item.isDisabled && c.item.isDisabled()) return true;
        return false;
    }

    window.ElixirCommandIndex = CommandIndex;
    return CommandIndex;
});

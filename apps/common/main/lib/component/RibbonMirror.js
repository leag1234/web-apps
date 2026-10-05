/*
 * SPDX-FileCopyrightText: 2026 ElixirOffice contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 *
 * Common.UI.RibbonMirror(src, opts): a ribbon button that mirrors another one.
 *
 * Office shows the same command in several places (Wrap Text in Layout and in
 * Picture Format, Page Break in Insert and Breaks in Layout...). Instead of
 * re-implementing a command, the twin follows the source button's state
 * (disabled, pressed, hint) and replays clicks and menu clicks on it, so the
 * existing controllers, locks and handlers keep doing the work. The source may
 * be hidden; it does not need to exist yet when the twin is created (pass a
 * function and call twin.mirrorAttach() later).
 *
 * src  : the source button, or a function returning it
 * opts : caption, iconCls, cls, split, hint (own tooltip),
 *        menu (a custom Common.UI.Menu: no cloning; see twin.clickSourceItem),
 *        click(src, e) (custom click handler), prepare(src), triggerShow (default true:
 *        the source menu 'show:before' runs before cloning, so its state is up to date)
 */

define([
    'core',
    'common/main/lib/component/Button',
    'common/main/lib/component/Menu'
], function () {
    'use strict';

    Common.UI.RibbonMirror = function (src, opts) {
        opts = opts || {};
        var getSrc = _.isFunction(src) ? src : function () { return src; },
            s = getSrc(),
            hasMenu = !!(opts.menu || (s ? s.menu : false)),
            twin = new Common.UI.Button({
                cls: opts.cls || 'btn-toolbar x-huge icon-top',
                iconCls: opts.iconCls || (s && s.options.iconCls) || '',
                caption: opts.caption !== undefined ? opts.caption : (s ? s.options.caption : ''),
                split: opts.split !== undefined ? opts.split : !!(s && s.options.split),
                menu: opts.menu ? opts.menu : (hasMenu ? new Common.UI.Menu({cls: (s && s.menu && s.menu.options && s.menu.options.cls) || '', items: []}) : false),
                enableToggle: !!(s && s.options.enableToggle),
                dataHint: '1',
                dataHintDirection: 'bottom',
                dataHintOffset: 'small'
            });
        twin.mirrorSource = getSrc;

        var attached = false;
        var attach = function () {
            var src = getSrc();
            if (attached || !src) return;
            attached = true;
            var twins = src._mirrorTwins || (src._mirrorTwins = []);
            twins.push(twin);
            if (twins.length === 1) {
                var origDisabled = src.setDisabled, origToggle = src.toggle, origHint = src.updateHint;
                src.setDisabled = function () {
                    origDisabled.apply(this, arguments);
                    var d = this.isDisabled();
                    _.each(this._mirrorTwins, function (t) { t.rendered && t.setDisabled(d); });
                };
                src.toggle = function (state, suppress) {
                    origToggle.apply(this, arguments);
                    var p = this.pressed;
                    _.each(this._mirrorTwins, function (t) { t.rendered && t.pressed !== p && t.toggle(p, true); });
                };
                src.updateHint = function (hint) {
                    origHint.apply(this, arguments);
                    _.each(this._mirrorTwins, function (t) { !t._ownHint && t.rendered && t.updateHint(hint); });
                };
            }
            twin.rendered && twin.setDisabled(src.isDisabled());
            twin.rendered && src.pressed && twin.toggle(true, true);
            if (opts.hint) { twin._ownHint = true; twin.rendered && twin.updateHint(opts.hint); }
            else if (src.options.hint && twin.rendered) twin.updateHint(src.options.hint);
        };
        twin.on('render:after', function () { attach(); var src = getSrc(); src && twin.setDisabled(src.isDisabled()); });
        twin.mirrorAttach = attach;

        // a click on the twin is a click on the source
        twin.on('click', function (btn, e) {
            var src = getSrc();
            if (!src || src.isDisabled()) return;
            if (twin.menu && !twin.split && !opts.click) return; // a plain menu button: the click only opens the menu
            if (opts.click) { opts.click(src, e); return; }
            if (src.options.enableToggle) src.toggle(twin.pressed);
            src.trigger('click', src, e);
        });

        // menus: rebuilt from the source menu each time they open
        var cloneItems = function (srcMenu, dstMenu) {
            dstMenu.removeAll();
            _.each(srcMenu.items, function (it) {
                if (it.caption === '--' || (it.options && it.options.caption === '--')) {
                    dstMenu.addItem({caption: '--'}); return;
                }
                if (!it.caption && !it.menu) return; // pickers and other custom items are not cloned
                var o = it.options || {};
                var cfg = {
                    caption: it.caption, iconCls: o.iconCls, value: it.value, checkable: it.checkable,
                    checkmark: o.checkmark, toggleGroup: o.toggleGroup ? 'mirror-' + o.toggleGroup : undefined,
                    checked: it.checked, disabled: it.disabled, allowDepress: o.allowDepress
                };
                var sub;
                if (it.menu) {
                    sub = new Common.UI.Menu({menuAlign: 'tl-tr', items: []});
                    cfg.menu = sub;
                }
                var mi = new Common.UI.MenuItem(cfg);
                mi.mirrorSource = it;
                dstMenu.addItem(mi);
                (it.$el && it.$el.css('display') === 'none') && mi.setVisible(false);
                if (sub) cloneItems(it.menu, sub);
            });
            dstMenu.off('item:click').on('item:click', function (menu, item, e) {
                var si = item.mirrorSource;
                if (!si || si.menu) return;
                // replays the click on the source item: checks it, fires its 'click' and the menu's 'item:click'
                if (srcMenu.rendered)
                    si.onItemClick({which: 1, target: {id: ''}, stopPropagation: function () {}, preventDefault: function () {}});
                else {
                    si.checkable && !(si.checked && si.toggleGroup) && si.setChecked(!si.checked);
                    srcMenu.trigger('item:click', srcMenu, si, e);
                }
            });
        };
        // a custom menu: its item values select the source menu item to click
        twin.clickSourceItem = function (value) {
            var src = getSrc(), si = src && src.menu && _.isObject(src.menu) && _.find(src.menu.items, function (it) { return it.value === value; });
            if (!si) return;
            if (src.menu.rendered)
                si.onItemClick({which: 1, target: {id: ''}, stopPropagation: function () {}, preventDefault: function () {}});
            else src.menu.trigger('item:click', src.menu, si, {});
        };
        if (hasMenu && !opts.menu) {
            twin.menu.on('show:before', function (menu) {
                var src = getSrc();
                if (!src || !src.menu || !_.isObject(src.menu)) return;
                opts.prepare && opts.prepare(src);
                if (opts.triggerShow !== false) src.menu.trigger('show:before', src.menu);
                cloneItems(src.menu, twin.menu);
            });
        }
        return twin;
    };

    return Common.UI.RibbonMirror;
});

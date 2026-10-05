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
 *  ElixirKeyTips.js (ElixirOffice)
 *
 *  Office KeyTips for the access keys of the ribbon (Alt, then H for Home,
 *  then a letter for a command), shared by the editors: the editors give their
 *  letters (tabs, commands, backstage) in a subclass (DE.Controllers.WordKeyTips,
 *  SSE.Controllers.ExcelKeyTips); a control without an Office equivalent keeps
 *  the letter that the hint manager assigns automatically. Level 0 shows the
 *  tabs and the quick access toolbar only, as Office does. Also: keyboard
 *  navigation inside the ribbon (Tab, arrows, Ctrl+arrows, Down, Esc,
 *  Ctrl+Shift+F10).
 */

define([
    'core'
], function () {
    'use strict';

    Common.Controllers = Common.Controllers || {};

    Common.Controllers.ElixirKeyTips = Backbone.Controller.extend({
        // editor tables: tab -> KeyTip, tab -> [[accessor, KeyTip]], backstage item id -> KeyTip
        tabs: {},
        commands: {},
        backstage: {},

        // accessor roots ("<root>.<path>", see resolve)
        roots: function () {
            var app = this.getApplication();
            return {
                tb: function() { var c = app.getController('Toolbar'); return c && c.toolbar; },
                review: function() { var c = app.getController('Common.Controllers.ReviewChanges'); return c && c.view; },
                viewtab: function() { var c = app.getController('ViewTab'); return c && c.view; },
                header: function() { var c = app.getController('Viewport'); return c && c.header; },
                ribbon: function() { var c = app.getController('RibbonTabs'); return c && c.view; }
            };
        },

        initialize: function() {
            Common.NotificationCenter.on('hints:before-show', _.bind(this.onBeforeShow, this));
            Common.NotificationCenter.on('app:ready', _.bind(this.onAppReady, this));
        },

        onAppReady: function() {
            var me = this;
            // level 0: the ribbon tabs and the quick access toolbar (not the side panes and the status bar)
            Common.UI.HintManager.setItemFilter && Common.UI.HintManager.setItemFilter(function(item, level) {
                if (level !== 0) return true;
                return item.closest('#toolbar, #header, .hedset, #app-title, #box-document-title').length > 0;
            });
            // Alt+letter held together opens the hints of a tab (Alt+F stays File, handled by the left menu)
            Common.UI.HintManager.setChordLetters && Common.UI.HintManager.setChordLetters(function() {
                // (not X: Alt+X converts Unicode codes; not on macOS, where Option+letter types characters)
                if (Common.Utils.isMac) return '';
                return _.filter(_.values(me.tabs), function(t) { return t.length === 1 && 'FAX'.indexOf(t) < 0; }).join('');
            });
            me.ready = true;
            // keyboard navigation inside the ribbon, as in Word
            document.addEventListener('keydown', _.bind(me.onRibbonKeyDown, me), true);
        },

        ribbonTabs: function() {
            return $('#toolbar .ribtab:visible > a[data-tab]');
        },

        ribbonControls: function() {
            var $panel = $('#toolbar .panel.active');
            return $panel.find('button, input, [tabindex="0"]').filter(function() {
                var $el = $(this);
                return $el.is(':visible') && !$el.prop('disabled') && !$el.closest('.disabled').length && $el.closest('.dropdown-menu').length === 0;
            });
        },

        focusEl: function($el) {
            if (!$el || !$el.length) return false;
            this.getApplication().getController('Viewport').getApi().asc_enableKeyEvents(false);
            $el[0].focus();
            return document.activeElement === $el[0];
        },

        // Tab, arrows, Ctrl+arrows, Down / Alt+Down, Esc and Ctrl+Shift+F10 in the ribbon
        onRibbonKeyDown: function(e) {
            var K0 = Common.UI.Keys;
            // KeyTips shown on the tabs: Tab, arrows leave them for the keyboard focus in the ribbon (Word)
            if (Common.UI.HintManager.isHintVisible() && !$('#file-menu-panel').is(':visible') &&
                    (e.keyCode === K0.TAB || e.keyCode === K0.LEFT || e.keyCode === K0.RIGHT || e.keyCode === K0.DOWN) &&
                    $('.hint-div:visible').filter(function() { return $(this).offset().top < 70; }).length > 0) {
                e.preventDefault(); e.stopPropagation();
                Common.UI.HintManager.clearHints(false, true);
                var $tab = $('#toolbar .ribtab.active > a');
                this.focusEl($tab);
                if (e.keyCode === K0.TAB || e.keyCode === K0.DOWN) this.focusEl(this.ribbonControls().first());
                return;
            }
            var active = document.activeElement;
            if (!active || !$(active).closest('#toolbar').length || Common.UI.HintManager.isHintVisible() || $(active).closest('.dropdown-menu').length)
                return;
            var K = Common.UI.Keys, me = this, api = this.getApplication().getController('Viewport').getApi();
            var $tabs = this.ribbonTabs(), tabIndex = $tabs.index(active);
            var stop = function() { e.preventDefault(); e.stopPropagation(); };
            var isInput = active.tagName === 'INPUT';

            if (e.keyCode === K.ESC) {
                stop();
                if (tabIndex < 0) me.focusEl($('#toolbar .ribtab.active > a'));
                else { api.asc_enableKeyEvents(true); Common.NotificationCenter.trigger('edit:complete'); }
                return;
            }
            if (e.keyCode === 121 && e.ctrlKey && e.shiftKey) {
                // Ctrl+Shift+F10: the ScreenTip of the focused control
                stop();
                var $t = $(active).closest('[data-original-title], .btn-group, .btn-slot, button');
                $t.trigger('mouseenter').trigger('mouseover');
                return;
            }
            if (tabIndex > -1) {
                if (e.keyCode === K.RIGHT || e.keyCode === K.LEFT) {
                    stop();
                    var next = $tabs.eq((tabIndex + (e.keyCode === K.RIGHT ? 1 : -1) + $tabs.length) % $tabs.length);
                    next.trigger('click');
                    _.defer(function() { me.focusEl(next); });
                } else if ((e.keyCode === K.DOWN || e.keyCode === K.TAB) && !e.shiftKey) {
                    stop();
                    me.focusEl(me.ribbonControls().first());
                }
                return;
            }
            var $controls = this.ribbonControls(), index = $controls.index(active);
            if (index < 0) return;
            if (e.keyCode === K.TAB || (!isInput && (e.keyCode === K.RIGHT || e.keyCode === K.LEFT) && !e.ctrlKey)) {
                stop();
                var back = e.keyCode === K.TAB ? e.shiftKey : e.keyCode === K.LEFT;
                if (back && index === 0) me.focusEl($('#toolbar .ribtab.active > a'));
                else me.focusEl($controls.eq((index + (back ? -1 : 1) + $controls.length) % $controls.length));
                return;
            }
            if ((e.keyCode === K.RIGHT || e.keyCode === K.LEFT) && e.ctrlKey) {
                // Ctrl+arrows: the first control of the next or previous group
                stop();
                var $groups = $('#toolbar .panel.active .group').filter(':visible');
                var g = $groups.index($(active).closest('.group'));
                var target = $groups.eq(g + (e.keyCode === K.RIGHT ? 1 : -1));
                if (g > -1 && target.length) {
                    var first = $controls.filter(function() { return target[0].contains(this); }).first();
                    me.focusEl(first);
                }
                return;
            }
            if (e.keyCode === K.DOWN && !isInput) {
                // Down (or Alt+Down): the list or the menu of the control
                var $group = $(active).closest('.btn-group');
                var $toggle = $group.find('.dropdown-toggle').filter(':visible').first();
                if ($toggle.length) {
                    stop();
                    $toggle.trigger('click');
                    _.defer(function() {
                        var $menu = $group.find('.dropdown-menu:visible').first();
                        $menu.length && $menu.focus();
                        $menu.find('a:visible').first().focus();
                    });
                }
            }
        },

        resolve: function(path) {
            var roots = this.roots();
            if (/^[#.\[]/.test(path))
                return $('#toolbar').find(path).first();
            var parts = path.split('.'), root = roots[parts[0]] && roots[parts[0]]();
            for (var i = 1; root && i < parts.length; i++)
                root = root[parts[i]];
            if (!root) return $();
            var $el = root.cmpEl || root.$el || (root.jquery ? root : null);
            return $el ? $el : $();
        },

        hintElement: function($el) {
            if (!$el || !$el.length) return $();
            if ($el.is('[data-hint]')) return $el;
            return $el.find('[data-hint]').first();
        },

        setTip: function($el, tip) {
            var $hint = this.hintElement($el);
            if (!$hint.length) return;
            $hint.attr('data-hint-title', tip).removeAttr('data-hint-title-lang');
        },

        onBeforeShow: function(level) {
            if (!this.ready) return;
            var me = this;
            if (level === 0) {
                $('#toolbar .ribtab > a[data-tab]').each(function() {
                    var tip = me.tabs[$(this).attr('data-tab')];
                    tip && $(this).attr('data-hint-title', tip).removeAttr('data-hint-title-lang');
                });
                // quick access toolbar: numbers from left to right, as in Word
                var n = 1;
                $('[data-hint="0"]').filter(':visible').filter(function() {
                    return $(this).closest('#toolbar').length === 0 && $(this).closest('#header, .hedset, #app-title, #box-document-title').length > 0 &&
                        /^(S|P|Z|Y|\d)$/.test($(this).attr('data-hint-title') || '');
                }).sort(function(a, b) { return a.getBoundingClientRect().left - b.getBoundingClientRect().left; }).each(function() {
                    // Word draws the quick access KeyTips under the buttons
                    $(this).attr({'data-hint-title': String(n++), 'data-hint-direction': 'bottom', 'data-hint-offset': '-6, 0'}).removeAttr('data-hint-title-lang');
                });
                // Search (Word: Alt, Q): the command search of the title bar, else the header search button
                var $search = $('#box-document-title .elx-search-input:visible');
                if ($search.length) {
                    $search.attr({'data-hint': '0', 'data-hint-title': 'Q', 'data-hint-direction': 'bottom', 'data-hint-offset': 'small'})
                        .removeAttr('data-hint-title-lang');
                } else
                    me.setTip(me.resolve('header.btnSearch'), 'Q');
                return;
            }
            if ($('#file-menu-panel').is(':visible')) {
                _.each(me.backstage, function(tip, id) {
                    var $el = $('#' + id);
                    me.setTip($el, tip);
                });
                return;
            }
            var tab = $('#toolbar .ribtab.active > a[data-tab]').attr('data-tab');
            _.each(me.commands[tab] || [], function(entry) {
                me.setTip(me.resolve(entry[0]), entry[1]);
            });
        }
    });
});

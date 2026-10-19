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
 *  WordShortcuts.js
 *
 *  Microsoft Word keyboard shortcuts that open interface elements (backstage
 *  pages, dialogs, panes). The editing shortcuts live in sdkjs
 *  (word/Editor/Shortcuts.js, word-shortcuts.js).
 *  Source: support.microsoft.com, "Keyboard shortcuts in Word" (docs/shortcuts-word.csv).
 */

define([
    'core',
    'documenteditor/main/app/view/GoToDialog'
], function () {
    'use strict';

    DE.Controllers.WordShortcuts = Backbone.Controller.extend({
        initialize: function() {
            Common.NotificationCenter.on('app:ready', _.bind(this.onAppReady, this));
        },

        onAppReady: function(mode) {
            this.mode = mode;
            this.api = this.getApplication().getController('Viewport').getApi();
            this.bindKeys();
        },

        bindKeys: function() {
            var me = this, mac = Common.Utils.isMac, keymap = {};
            var bind = function(keys, fn) {
                keymap[keys] = function(e) {
                    if (!me.api || !me.mode) return true;
                    var r = fn.call(me, e);
                    return r === true;   // false: handled, prevent the browser default
                };
            };

            // File
            bind(mac ? 'command+o' : 'ctrl+o,ctrl+f12,ctrl+alt+f2', this.onOpen);
            bind(mac ? 'command+n' : 'ctrl+n', this.onNew);
            bind('shift+f12,alt+shift+f2', this.onSave);
            !mac && bind('f12', this.onSaveAs);   // macOS Cmd+Shift+S: LeftMenu
            bind('ctrl+alt+i,ctrl+f2,ctrl+shift+f12', this.onPrint);
            !mac && bind('alt+b', this.onRename);
            // Windows, panes, ribbon
            bind('f6', function() { return this.cycleRegions(1); });
            bind('shift+f6', function() { return this.cycleRegions(-1); });
            bind(mac ? 'command+alt+r' : 'ctrl+f1', this.onToggleRibbon);
            bind('alt+shift+c', this.onClosePane);
            bind('alt+f8,alt+f11', this.onMacros);
            !mac && bind('alt+shift+f10', this.onActionsMenu);
            !mac && bind('ctrl+alt+p', function() { return false; });   // Print Layout: the only layout view
            // Dialogs
            bind(mac ? 'command+d' : 'ctrl+d,ctrl+shift+f', this.onFontDialog);
            bind(mac ? 'command+alt+g,f5' : 'ctrl+g,f5', this.onGoTo);
            bind(mac ? 'command+shift+f5' : 'ctrl+shift+f5', this.onBookmarks);

            Common.util.Shortcuts.delegateShortcuts({shortcuts: keymap});
        },

        isModal: function() {
            return Common.Utils.ModalWindow.isVisible();
        },

        leftMenu: function() {
            return this.getApplication().getController('LeftMenu');
        },

        showBackstage: function(panel) {
            var lm = this.leftMenu();
            if (!lm || !lm.leftMenu) return;
            Common.UI.Menu.Manager.hideAll();
            lm.leftMenu.showMenu(panel ? 'file:' + panel : 'file');
        },

        onOpen: function() {
            if (this.isModal()) return true;
            // File > Open in Word; here the recent files when the integration provides them
            this.showBackstage(this.mode.canOpenRecent ? 'recent' : undefined);
            return false;
        },

        onNew: function() {
            if (this.isModal()) return true;
            this.mode.canCreateNew && this.showBackstage('new');
            return false;
        },

        onSave: function() {
            if (this.isModal()) return true;
            this.api.asc_Save && this.api.asc_Save(false);
            return false;
        },

        onSaveAs: function() {
            if (this.isModal()) return true;
            var lm = this.leftMenu();
            lm && lm.onShortcut('save');
            return false;
        },

        onPrint: function() {
            if (this.isModal()) return true;
            this.mode.canPrint && Common.NotificationCenter.trigger('file:print');
            return false;
        },

        onRename: function() {
            if (this.isModal() || !this.mode.canRename) return true;
            var lm = this.leftMenu();
            lm && lm.clickMenuFileItem(null, 'rename');
            return false;
        },

        onFontDialog: function() {
            if (this.isModal() || !this.mode.isEdit) return true;
            // same as the dialog launcher of the Font group (Word's Font dialog of WordFormatting)
            var fmt = this.getApplication().getController('WordFormatting');
            if (fmt && fmt.onGroupLauncher) fmt.onGroupLauncher('font');
            else {
                var toolbar = this.getApplication().getController('Toolbar');
                toolbar && toolbar.onGroupLauncher && toolbar.onGroupLauncher('font');
            }
            return false;
        },

        onGoTo: function() {
            if (this.isModal()) return true;
            var me = this;
            Common.UI.Menu.Manager.hideAll();
            this.gotoDialog = new DE.Views.GoToDialog({
                api: this.api,
                handler: function() { me.api.asc_enableKeyEvents(true); }
            });
            this.gotoDialog.on('close', function() {
                me.gotoDialog = null;
                Common.NotificationCenter.trigger('edit:complete');
            });
            this.gotoDialog.show();
            return false;
        },

        onBookmarks: function() {
            if (this.isModal() || !this.mode.isEdit) return true;
            var links = this.getApplication().getController('Links');
            links && links.onBookmarksClick && links.onBookmarksClick();
            return false;
        },

        onClosePane: function() {
            if (this.isModal()) return true;
            var lm = this.leftMenu();
            if (lm && lm.leftMenu && lm.leftMenu.btnComments && lm.leftMenu.btnComments.pressed) {
                lm.leftMenu.close();
                Common.NotificationCenter.trigger('layout:changed', 'leftmenu');
                Common.NotificationCenter.trigger('edit:complete');
            }
            return false;
        },

        onToggleRibbon: function() {
            if (this.isModal()) return true;
            var toolbar = this.getApplication().getController('Toolbar');
            if (!toolbar || !toolbar.toolbar) return true;
            // the ribbon controller also updates its pin button
            var ribbon = this.getApplication().getController('RibbonTabs');
            if (ribbon && ribbon.toggleRibbon && ribbon.toolbar) ribbon.toggleRibbon();
            else toolbar.onChangeCompactView(null, !toolbar.toolbar.isCompact());
            return false;
        },

        onMacros: function() {
            if (this.isModal() || !this.mode.isEdit || !Common.Views.MacrosDialog) return true;
            var view = this.getApplication().getController('ViewTab');
            if (view && view.onClickMacros) view.onClickMacros();
            else (new Common.Views.MacrosDialog({api: this.api})).show();
            return false;
        },

        onActionsMenu: function() {
            // Alt+Shift+F10: the paste options after a paste
            var dh = this.getApplication().getController('DocumentHolder');
            var btn = dh && dh.btnSpecialPaste;
            if (btn && btn.cmpEl && btn.cmpEl.is(':visible') && btn.menu) {
                !btn.menu.isVisible() && $('button', btn.cmpEl).click();
                return false;
            }
            return true;
        },

        // F6 / Shift+F6: document, task pane, status bar, ribbon
        cycleRegions: function(dir) {
            if (this.isModal()) return true;
            var regions = [
                {name: 'document', el: null},
                {name: 'layoutoptions', el: $('.elx-lo-btn:visible, .elx-lo-pane:visible')},
                {name: 'pane', el: $('#left-panel .left-panel:visible, #left-menu .left-panel:visible, #right-menu:visible').first()},
                {name: 'statusbar', el: $('#statusbar')},
                {name: 'ribbon', el: $('#toolbar')}
            ].filter(function(r) { return r.name === 'document' || (r.el && r.el.length && r.el.is(':visible')); });

            var active = document.activeElement, current = 0;
            for (var i = 1; i < regions.length; i++) {
                if (active && regions[i].el.filter(function () { return this.contains(active); }).length) { current = i; break; }
            }
            // next region that can take the focus (a visible pane may have nothing focusable)
            for (var step = 1; step <= regions.length; step++) {
                var next = regions[(current + dir * step + regions.length * step) % regions.length];
                if (next.name === 'document') {
                    this.api.asc_enableKeyEvents(true);
                    Common.NotificationCenter.trigger('edit:complete');
                    return false;
                }
                var target = next.name === 'layoutoptions' ? next.el.filter('.elx-lo-btn').first() : next.name === 'ribbon' ? next.el.find('.ribtab.active > a, li.active > a').filter(':visible').first() : $();
                if (!target.length)
                    target = next.el.find('button:visible:not([disabled]), a[href]:visible, input:visible:not([disabled])').first();
                if (target.length) {
                    this.api.asc_enableKeyEvents(next.name === 'layoutoptions');
                    target[0].focus();
                    if (document.activeElement === target[0])
                        return false;
                }
            }
            return false;
        }
    });
});

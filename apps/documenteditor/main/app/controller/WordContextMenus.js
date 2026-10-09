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
 *  WordContextMenus.js
 *
 *  The right-click menus of text, tables and pictures in Word's order:
 *  spelling suggestions first, then Cut / Copy / Paste, the commands of the
 *  object (fields, table cells, picture), Font... and Paragraph..., Link and
 *  New Comment last; Word's captions, Fluent icons and the shortcut of each
 *  command on the right. The items, their visibility rules and their actions
 *  stay those of DocumentHolder; only their order, captions and separators
 *  change.
 */

define([
    'core'
], function () {
    'use strict';

    // Languages with translated captions: elsewhere the editor's own captions are kept
    var LANGS = ['en', 'fr', 'de', 'es', 'it', 'pt', 'pt-pt', 'nl', 'ru', 'pl', 'ja', 'zh'];

    // [menu, groups]; an entry is a DocumentHolder view property, or @name for an item found by rule()
    var ORDER = {
        textMenu: [
            ['menuSpellPara', 'menuSpellMorePara'],
            ['menuIgnoreSpellPara', 'menuIgnoreAllSpellPara', 'menuToDictionaryPara', 'langParaMenu'],
            ['menuParaCut', 'menuParaCopy', 'menuParaPaste'],
            ['@elixir'],
            ['menuParaAccept', 'menuParaReject'],
            ['menuParaRefreshField', 'menuParaEditField', 'menuParaFieldCodes', 'menuParaTOCRefresh', 'menuParaTOCSettings',
                'menuParaRemoveControl', 'menuParaControlSettings', 'menuEquationInsertCaption'],
            ['@font', 'menuParagraphAdvanced', 'menuParaStartNewList', 'menuParaContinueNumbering', 'menuParaStartNumberingFrom',
                'menuParaListIndents', 'menuParagraphVAlign', 'menuParagraphDirection', 'menuFrameAdvanced', 'menuDropCapAdvanced', 'menuParagraphEquation'],
            ['menuAddHyperlinkPara', '@hyperlink', 'menuParaFollow'],
            ['menuAddCommentPara'],
            ['@style', 'menuParaPrint', 'menuParaAssistant']
        ],
        tableMenu: [
            ['menuSpellTable', 'menuSpellMoreTable'],
            ['menuIgnoreAllSpellTable', 'menuToDictionaryTable', 'langTableMenu'],
            ['menuTableCut', 'menuTableCopy', 'menuTablePaste'],
            ['@elixir'],
            ['menuTableAccept', 'menuTableReject'],
            ['menuTableRefreshField', 'menuTableEditField', 'menuTableFieldCodes'],
            ['menuTableInsertText', 'menuTableDeleteText', 'menuTableSelectText', 'mnuTableMerge', 'mnuTableSplit',
                'menuTableDistRows', 'menuTableDistCols', 'menuTableDirection', 'menuTableCellAlign', 'menuTableAdvanced'],
            ['menuTableInsertCaption', '@font', 'menuParagraphAdvancedInTable', '@bullets', 'menuTableEquation'],
            ['menuAddHyperlinkTable', '@hyperlink', 'menuTableFollow'],
            ['menuAddCommentTable'],
            ['menuTableStretchContentControl', 'menuTableRemoveForm', '@control', 'menuTableTOC', 'menuStyleInTable', 'menuTablePrint']
        ],
        pictureMenu: [
            ['menuImgCut', 'menuImgCopy', 'menuImgPaste'],
            ['@elixir'],
            ['menuEditObject', 'menuChartEdit', 'menuImgEditPoints'],
            ['menuImgAccept', 'menuImgReject'],
            ['menuSignatureEditSign', 'menuSignatureEditSetup'],
            ['menuImgRemoveControl', 'menuImgControlSettings'],
            ['menuImgReplace', 'menuImgCrop', 'menuImgResetCrop', 'menuOriginalSize'],
            ['menuImageArrange', 'menuImageAlign', 'menuShapesMerge', 'menuImgRotate'],
            ['menuAddHyperlinkPic', '@hyperlink'],
            ['menuSaveAsPicture', 'menuInsertCaption', 'menuImageWrap', 'menuImgStretchContentControl', 'menuImageAdvanced'],
            ['menuImgPrint']
        ]
    };

    // Word's captions and icons (the string keys are those of this controller in the locale files)
    var LOOK = {
        menuAddCommentPara: ['textNewComment'], menuAddCommentTable: ['textNewComment'],
        menuParagraphAdvanced: ['textParagraph', 'menu__icon btn-paragraph'],
        menuParagraphAdvancedInTable: ['textParagraph', 'menu__icon btn-paragraph'],
        menuEquationInsertCaption: ['textInsertCaption', 'menu__icon btn-menu-caption'],
        menuTableInsertCaption: ['textInsertCaption', 'menu__icon btn-menu-caption'],
        menuInsertCaption: ['textInsertCaption', 'menu__icon btn-menu-caption'],
        menuTableAdvanced: ['textTableProperties', 'menu__icon btn-menu-table-props'],
        mnuTableMerge: ['textMergeCells', 'menu__icon btn-menu-merge-cells'],
        mnuTableSplit: ['textSplitCells', 'menu__icon btn-menu-split-cells'],
        menuTableInsertText: ['textInsert', 'menu__icon btn-menu-insert-cells'],
        menuTableDeleteText: ['textDelete', 'menu__icon btn-menu-delete-cells'],
        menuTableCellAlign: ['textCellAlignment'],
        menuImageAdvanced: [null, 'menu__icon btn-menu-size-position'],   // caption set on show (pictures only)
        menuImageWrap: ['textWrapText'],
        menuImgReplace: ['textChangePicture', 'menu__icon btn-menu-change-picture'],
        menuSaveAsPicture: ['textSaveAsPicture', 'menu__icon btn-menu-save-picture']
    };

    // shortcuts shown on the right: Windows/Linux, macOS
    var KEYS = {
        menuParaCut: ['Ctrl+X', '⌘X'], menuTableCut: ['Ctrl+X', '⌘X'], menuImgCut: ['Ctrl+X', '⌘X'],
        menuParaCopy: ['Ctrl+C', '⌘C'], menuTableCopy: ['Ctrl+C', '⌘C'], menuImgCopy: ['Ctrl+C', '⌘C'],
        menuParaPaste: ['Ctrl+V', '⌘V'], menuTablePaste: ['Ctrl+V', '⌘V'], menuImgPaste: ['Ctrl+V', '⌘V'],
        '@font': ['Ctrl+D', '⌘D'],
        menuAddHyperlinkPara: ['Ctrl+K', '⌘K'], menuAddHyperlinkTable: ['Ctrl+K', '⌘K'], menuAddHyperlinkPic: ['Ctrl+K', '⌘K'],
        menuAddCommentPara: ['Ctrl+Alt+M', '⌥⌘A'], menuAddCommentTable: ['Ctrl+Alt+M', '⌥⌘A'],
        menuParaRefreshField: ['F9', 'F9'], menuTableRefreshField: ['F9', 'F9'],
        menuParaFieldCodes: ['Shift+F9', '⇧F9'], menuTableFieldCodes: ['Shift+F9', '⇧F9'],
        menuParaPrint: ['Ctrl+P', '⌘P'], menuTablePrint: ['Ctrl+P', '⌘P'], menuImgPrint: ['Ctrl+P', '⌘P']
    };

    DE.Controllers.WordContextMenus = Backbone.Controller.extend(_.extend({
        initialize: function() {
            Common.NotificationCenter.on('app:ready', _.bind(this.onAppReady, this));
        },

        onAppReady: function(mode) {
            this.mode = mode;
            var me = this, tries = 0;
            var wait = function() {
                var dh = me.getApplication().getController('DocumentHolder'), view = dh && dh.documentHolder;
                if (view && view.textMenu && view.tableMenu && view.pictureMenu) {
                    me.setup(view);
                } else if (++tries < 60) {
                    setTimeout(wait, 500);
                }
            };
            wait();
        },

        lang: function() {
            var l = ((this.mode && this.mode.lang) || 'en').toLowerCase();
            return LANGS.indexOf(l) > -1 ? l : (LANGS.indexOf(l.split(/[-_]/)[0]) > -1 ? l.split(/[-_]/)[0] : null);
        },

        // items of a menu that are not view properties
        rule: function(view, menu, name) {
            var named = _.values(view);
            return _.find(menu.items, function(it) {
                if (named.indexOf(it) > -1) return false;
                switch (name) {
                    case 'hyperlink': return !!it.menu && /btn-inserthyperlink/.test(it.iconCls || (it.options && it.options.iconCls) || '');
                    case 'style': return !!it.menu && it.caption === view.styleText;
                    case 'control': return !!it.menu && it.caption === view.textContentControls;
                    case 'bullets': return !!it.menu && it.caption === view.bulletsText;
                }
                return false;
            });
        },

        setup: function(view) {
            var me = this;
            this.view = view;
            // spelling suggestions arrive after the menu is open: mark them when they are added
            var addWordVariants = view.addWordVariants;
            view.addWordVariants = function() {
                var r = addWordVariants.apply(this, arguments);
                _.each([view.textMenu, view.tableMenu], function(menu) {
                    menu && _.each(menu.items, function(it) {
                        it.options && it.options.spellword && it.cmpEl && it.cmpEl.addClass('spell-suggestion');
                    });
                });
                view.textMenu && view.textMenu._wordSeparators && me.updateSeparators(view.textMenu);
                return r;
            };
            this.fontItems = {};
            _.each(ORDER, function(groups, menuName) {
                var menu = view[menuName];
                if (!menu || menu.rendered) return;    // reordered before the first rendering only
                var placed = [], order = [], seps = [];
                var groupItems = _.map(groups, function(group) {
                    return _.compact(_.map(group, function(name) {
                        var item;
                        if (name === '@font') {
                            item = me.fontItems[menuName] = new Common.UI.MenuItem({
                                caption: me.textFont,
                                iconCls: 'menu__icon btn-menu-font'
                            });
                            item.on('click', function() { me.onFont(); });
                        } else if (name === '@elixir') {
                            // Elixir AI: opens the "Ask Elixir" menu where the context menu was
                            var AI = window.Common && Common.ElixirAI;
                            item = new Common.UI.MenuItem({
                                caption: AI && AI.t ? AI.t('menuAsk') : 'Ask Elixir',
                                iconCls: 'menu__icon elixir-ai-icon elixir-ai-icon-sparkle'
                            });
                            me.elixirItems = me.elixirItems || [];
                            me.elixirItems.push(item);
                            item.on('click', function() { me.onAskElixir(menu); });
                        } else if (name.charAt(0) === '@') {
                            item = me.rule(view, menu, name.substring(1));
                        } else {
                            item = view[name];
                        }
                        if (!item || placed.indexOf(item) > -1 || (name.charAt(0) !== '@' && menu.items.indexOf(item) < 0)) return null;
                        placed.push(item);
                        item._wordKey = name;
                        return item;
                    }));
                });
                _.each(groupItems, function(items, i) {
                    if (!items.length) return;
                    if (order.length) {
                        var sep = new Common.UI.MenuItem({caption: '--'});
                        sep._wordSeparator = true;
                        seps.push({sep: sep, before: items});
                        order.push(sep);
                    }
                    order.push.apply(order, items);
                });
                // the others (kept in their order) and the editor's separators, hidden
                _.each(menu.items, function(it) {
                    if (placed.indexOf(it) > -1) return;
                    if (it.caption === '--') { it._wordHidden = true; order.push(it); return; }
                    order.push(it);
                });
                menu.items = order;
                menu._wordSeparators = seps;
                me.applyLook(menu, view);
                // after DocumentHolder has set the visibility of the items (initMenu runs once the menu is shown)
                var initMenu = menu.options.initMenu;
                menu.options.initMenu = function(value) {
                    initMenu && initMenu.apply(this, arguments);
                    try { me.beforeShow(menu, view); me.afterShow(menu); } catch (e) { console.warn('WordContextMenus', e); }
                };
            });
        },

        applyLook: function(menu, view) {
            var me = this, lang = this.lang();
            _.each(menu.items, function(it) {
                var look = LOOK[it._wordKey];
                if (!look) return;
                if (lang && look[0] && me[look[0]]) {
                    it.caption = me[look[0]];
                    it.options && (it.options.caption = it.caption);
                }
                if (look[1]) {
                    it.iconCls = look[1];
                    it.options && (it.options.iconCls = look[1]);
                }
            });
        },

        beforeShow: function(menu, view) {
            var me = this;
            // Word has these in the Paragraph dialog only
            view.menuParagraphBreakBefore && view.menuParagraphBreakBefore.setVisible(false);
            view.menuParagraphKeepLines && view.menuParagraphKeepLines.setVisible(false);
            // Font... with the paragraph commands
            var font = this.fontItems[_.findKey(ORDER, function(g, name) { return view[name] === menu; })];
            if (font) {
                var anchor = menu === view.textMenu ? view.menuParagraphAdvanced : view.menuParagraphAdvancedInTable;
                font.setVisible(!!(anchor && anchor.el && anchor.el.style.display !== 'none') && !!this.mode.isEdit);
            }
            _.each(menu.items, function(it) { it._wordHidden && it.setVisible(false); });
            var ai = this.getApplication().getController('Common.Controllers.ElixirAI'), aiOn = !!(ai && ai.isActive && ai.isActive());
            _.each(this.elixirItems || [], function(it) { it.setVisible(aiOn); });
            this.updateSeparators(menu);
        },

        onAskElixir: function(menu) {
            var $m = menu.cmpEl.find('> .dropdown-menu').addBack('.dropdown-menu').first();
            if (!$m.length) $m = menu.cmpEl;
            var off = $m.offset() || {left: 100, top: 100};
            var ep = Common.ElixirAI && Common.ElixirAI.entryPoints && Common.ElixirAI.entryPoints.word;
            _.defer(function() { ep && ep.openAskMenu && ep.openAskMenu(off.left, off.top); });
        },

        afterShow: function(menu) {
            var view = this.view, mac = Common.Utils.isMac, me = this;
            _.each(menu.items, function(it) {
                // the editor shows its own separators again while opening the menu
                it._wordHidden && it.setVisible(false);
                // spelling suggestions in bold, as in Word
                it.options && it.options.spellword && it.cmpEl && it.cmpEl.addClass('spell-suggestion');
                // shortcuts on the right
                var keys = KEYS[it._wordKey];
                if (keys && it.cmpEl) {
                    var $a = it.cmpEl.find('> a').first();
                    if ($a.length && !$a.find('.menu-item-shortcut').length)
                        $a.append('<span class="menu-item-shortcut">' + keys[mac ? 1 : 0] + '</span>');
                }
            });
            // Word: "Size and Position..." for a picture (the editor sets the caption when the menu opens)
            if (menu === view.pictureMenu && view.menuImageAdvanced.caption === view.imageText && this.lang())
                view.menuImageAdvanced.setCaption(this.textSizeAndPosition);
            this.updateSeparators(menu);
        },

        // a separator only between two visible groups
        updateSeparators: function(menu) {
            // the item's own display, whether the menu is open yet or not
            var visible = function(it) { return !!it && (it.rendered && it.el ? it.el.style.display !== 'none' : it.visible !== false); };
            var seen = false;
            _.each(menu._wordSeparators || [], function(s) {
                var before = menu.items.slice(0, menu.items.indexOf(s.sep));
                var hasBefore = _.some(before, function(it) { return it.caption !== '--' && !it._wordHidden && visible(it); });
                var hasAfter = _.some(s.before, visible);
                s.sep.setVisible(hasBefore && hasAfter);
            });
        },

        onFont: function() {
            var fmt = this.getApplication().getController('WordFormatting');
            if (fmt && fmt.onGroupLauncher) { fmt.onGroupLauncher('font'); return; }
            var toolbar = this.getApplication().getController('Toolbar');
            toolbar && toolbar.onGroupLauncher && toolbar.onGroupLauncher('font');
        },

        textFont: 'Font...',
        textParagraph: 'Paragraph...',
        textNewComment: 'New Comment',
        textTableProperties: 'Table Properties...',
        textMergeCells: 'Merge Cells',
        textSplitCells: 'Split Cells...',
        textInsertCaption: 'Insert Caption...',
        textSizeAndPosition: 'Size and Position...',
        textWrapText: 'Wrap Text',
        textChangePicture: 'Change Picture',
        textSaveAsPicture: 'Save as Picture...',
        textCellAlignment: 'Cell Alignment',
        textInsert: 'Insert',
        textDelete: 'Delete'

    }, DE.Controllers.WordContextMenus || {}));
});

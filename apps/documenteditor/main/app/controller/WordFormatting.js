/*
 * SPDX-FileCopyrightText: 2026 ElixirOffice contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 *
 * DE.Controllers.WordFormatting: Word-like formatting on the Home tab of the document editor.
 *  - Styles gallery: Word's list and order, "AaBbCcDd" previews with the style name below, live
 *    preview on hover, Create a Style / Clear Formatting / Apply Styles in the expanded gallery,
 *    Modify... in the item menu;
 *  - Styles pane (Ctrl+Alt+Shift+S and the Styles group launcher), Apply Styles box,
 *    Modify Style dialog.
 * It only listens to the Toolbar view and the api; the hooks are its declaration in app.js and
 * app_dev.js.
 */

define([
    'core',
    'documenteditor/main/app/view/StylesPane',
    'documenteditor/main/app/view/FontDialog',
    'documenteditor/main/app/view/WordParagraphDialog'
], function () {
    'use strict';

    // size of a gallery preview (CSS px); the item adds the style name below it
    var GALLERY_PREVIEW_W = 84,
        GALLERY_PREVIEW_H = 40;

    DE.Controllers.WordFormatting = Backbone.Controller.extend(_.extend({
        initialize: function () {
            var me = this;
            this._preview = {timer: 0, endTimer: 0, active: false, name: null};
            this.addListeners({
                'Toolbar': {
                    'render:before': function (toolbar) { me.onToolbarRenderBefore(toolbar); },
                    'render:after': function (toolbar) { me.onToolbarRenderAfter(toolbar); },
                    'group:launcher': function (group) { me.onGroupLauncher(group); }
                }
            });
            Common.NotificationCenter.on('app:ready', this.onAppReady.bind(this));
        },

        onLaunch: function () {},

        getApi: function () {
            var main = this.getApplication().getController('Main');
            return main && main.api;
        },

        appOptions: function () {
            var main = this.getApplication().getController('Main');
            return (main && main.appOptions) || {};
        },

        canEdit: function () {
            var o = this.appOptions();
            return !!(o.isEdit && !o.isEditMailMerge && !o.isEditDiagram && !o.isEditOle);
        },

        translate: function (name) {
            var main = this.getApplication().getController('Main');
            return (main && main.translationTable && main.translationTable[name]) || name;
        },

        // ---------------------------------------------------------------- toolbar wiring
        onToolbarRenderBefore: function (toolbar) {
            this.toolbar = toolbar;
            var api = this.getApi();
            if (api && api.asc_SetStyleGalleryOptions)
                api.asc_SetStyleGalleryOptions({wordLike: true, previewText: 'AaBbCcDd', width: GALLERY_PREVIEW_W, height: GALLERY_PREVIEW_H});
        },

        onToolbarRenderAfter: function (toolbar) {
            this.toolbar = toolbar;
            if (!this.canEdit() || !toolbar.listStyles) return;
            this.bindGalleryPreview(toolbar.listStyles);
            this.extendGalleryMenus(toolbar);
        },

        onAppReady: function () {
            var api = this.getApi(), me = this;
            if (!api || !this.canEdit()) return;
            this.api = api;
            api.asc_UpdateStyleGallery && api.asc_UpdateStyleGallery();
            api.asc_registerCallback('asc_onParaStyleName', function (name) {
                me._currentStyle = name;
                me.stylesPane && me.stylesPane.isVisible() && me.stylesPane.setCurrent(name);
                me.applyStylesDlg && me.applyStylesDlg.isVisible() && me.applyStylesDlg.setCurrent(name);
                me.checkHeadingLevels(name);
            });
            api.asc_registerCallback('asc_onInitEditorStyles', function () {
                me.stylesPane && me.stylesPane.isVisible() && me.refreshStylesPane();
            });
            var keymap = {};
            // Word: Ctrl+D opens the Font dialog
            keymap['command+d,ctrl+d'] = function (e) {
                if (!me.isEditingAllowed() || $('.asc-window.modal:visible').length) return true;
                me.openFontDialog();
                e && e.preventDefault && e.preventDefault();
                return false;
            };
            keymap['command+alt+shift+s,ctrl+alt+shift+s'] = function (e) {
                if (!me.isEditingAllowed()) return true;
                me.toggleStylesPane();
                e && e.preventDefault && e.preventDefault();
                return false;
            };
            Common.util.Shortcuts.delegateShortcuts({shortcuts: keymap});
        },

        isEditingAllowed: function () {
            return this.canEdit() && this.toolbar && this.toolbar.listStyles && !this.toolbar.listStyles.isDisabled();
        },

        // Word shows Heading N+1 in the gallery once Heading N is used: regenerate when a heading
        // level appears that the gallery does not show the next level for.
        checkHeadingLevels: function (name) {
            var m = /^heading ([1-8])$/i.exec(name || '');
            if (!m || !this.toolbar || !this.toolbar.listStyles || !this.api) return;
            var next = 'Heading ' + (parseInt(m[1]) + 1),
                store = this.toolbar.listStyles.menuPicker.store;
            if (store.length && !store.find(function (r) { return (r.get('title') || '').toLowerCase() === next.toLowerCase(); })) {
                var me = this;
                clearTimeout(this._regenTimer);
                this._regenTimer = setTimeout(function () { me.api.asc_UpdateStyleGallery && me.api.asc_UpdateStyleGallery(); }, 300);
            }
        },

        onGroupLauncher: function (group) {
            _.isArray(group) && (group = group[0]);
            if (!this.isEditingAllowed()) return;
            switch (group) {
                case 'styles': this.toggleStylesPane(true); break;
                case 'font': this.openFontDialog(); break;
                case 'paragraph': this.openParagraphDialog(); break;
            }
        },

        // ---------------------------------------------------------------- Font dialog
        openFontDialog: function () {
            var api = this.api || this.getApi(), me = this;
            if (!api || !api.asc_GetFontDialogProps) return;
            this.endPreview();
            var text = '';
            try { text = (api.asc_GetSelectedText(true) || '').replace(/\s+/g, ' ').trim().substr(0, 40); } catch (e) {}
            (new DE.Views.FontDialog({
                api: api,
                props: api.asc_GetFontDialogProps(),
                fontStore: this.getFontStore(),
                previewText: text,
                handler: function (result, changes) {
                    result === 'ok' && !_.isEmpty(changes) && api.asc_SetFontDialogProps(changes);
                    Common.NotificationCenter.trigger('edit:complete', me.toolbar);
                }
            })).show();
        },

        // ---------------------------------------------------------------- Paragraph and Tabs dialogs
        getParagraphProps: function () {
            var api = this.api, elements = api ? api.getSelectedElements() : null;
            for (var i = elements ? elements.length - 1 : -1; i >= 0; i--) {
                if (elements[i].get_ObjectType() === Asc.c_oAscTypeSelectElement.Paragraph)
                    return elements[i].get_ObjectValue();
            }
            return null;
        },

        openParagraphDialog: function (View) {
            var api = this.api || this.getApi(), me = this,
                props = this.getParagraphProps(),
                dh = this.getApplication().getController('DocumentHolder');
            if (!props) return;
            var win = new (View || DE.Views.WordParagraphDialog)({
                tableStylerRows: 2,
                tableStylerColumns: 1,
                paragraphProps: props,
                borderProps: dh && dh.borderAdvancedProps,
                isChart: false,
                isSmartArtInternal: false,
                api: api,
                handler: function (result, value) {
                    if (result === 'tabs') {
                        // Word applies the Paragraph dialog, then opens the Tabs dialog
                        value = this.getSettings();
                    }
                    if ((result === 'ok' || result === 'tabs') && value) {
                        dh && (dh.borderAdvancedProps = value.borderProps);
                        api.paraApply(value.paragraphProps);
                    }
                    if (result === 'tabs')
                        _.defer(function () { me.openParagraphDialog(DE.Views.WordTabsDialog); });
                    else
                        Common.NotificationCenter.trigger('edit:complete', me.toolbar);
                }
            });
            win.show();
            return win;
        },

        // ---------------------------------------------------------------- live preview
        bindGalleryPreview: function (ls) {
            var me = this;
            var recordOf = function (el) {
                var id = $(el).find('.style').attr('id');
                if (!id) return null;
                return ls.menuPicker.store.findWhere({id: id}) || (ls.fieldPicker.store && ls.fieldPicker.store.findWhere({id: id}));
            };
            ls.cmpEl.on('mouseenter', '.item', function (e) {
                if (ls.isDisabled() || me._preview.suspended) return;
                var rec = recordOf(e.currentTarget);
                rec && me.schedulePreview(rec.get('title'));
            });
            ls.cmpEl.on('mouseleave', '.item', function () { me.scheduleEndPreview(); });
            ls.cmpEl.on('mouseleave', function () { me.scheduleEndPreview(); });
            // before the click applies the style for real
            ls.cmpEl.on('mousedown', function () { me.endPreview(); });
            ls.openButton.menu.on('hide:before', function () { me.endPreview(); });
            ls.on('contextmenu', function () { me.endPreview(); });
            // keyboard, wheel or a click anywhere end it too
            var stop = function () { me._preview.active && me.endPreview(); };
            document.addEventListener('keydown', stop, true);
            document.addEventListener('wheel', stop, true);
        },

        schedulePreview: function (name) {
            var me = this, p = this._preview;
            clearTimeout(p.timer);
            clearTimeout(p.endTimer);
            p.timer = setTimeout(function () { me.startPreview(name); }, p.active ? 80 : 300);
        },

        scheduleEndPreview: function () {
            var me = this, p = this._preview;
            clearTimeout(p.timer);
            clearTimeout(p.endTimer);
            p.endTimer = setTimeout(function () { me.endPreview(); }, 120);
        },

        startPreview: function (name) {
            var api = this.api || this.getApi(), p = this._preview;
            if (!api || !api.asc_StartStylePreview) return;
            if (p.active && p.name === name) return;
            p.active = !!api.asc_StartStylePreview(name);
            p.name = p.active ? name : null;
        },

        endPreview: function () {
            var api = this.api || this.getApi(), p = this._preview;
            clearTimeout(p.timer);
            clearTimeout(p.endTimer);
            if (p.active) {
                p.active = false;
                p.name = null;
                api && api.asc_EndStylePreview && api.asc_EndStylePreview();
            }
        },

        // ---------------------------------------------------------------- gallery menus
        extendGalleryMenus: function (toolbar) {
            var me = this, ls = toolbar.listStyles, menu = ls.openButton.menu;
            // expanded gallery: Create a Style / Clear Formatting / Apply Styles...
            toolbar.listStylesAdditionalMenuItem && toolbar.listStylesAdditionalMenuItem.setCaption(this.textCreateStyle);
            toolbar.listStylesAdditionalMenuItem && toolbar.listStylesAdditionalMenuItem.setIconCls && toolbar.listStylesAdditionalMenuItem.setIconCls('menu__icon btn-style-new');
            this.mnuGalleryClear = new Common.UI.MenuItem({caption: this.textClearFormatting, iconCls: 'menu__icon btn-clearstyle'});
            this.mnuGalleryApply = new Common.UI.MenuItem({caption: this.textApplyStyles, iconCls: 'menu__icon btn-style-apply'});
            menu.addItem(this.mnuGalleryClear);
            menu.addItem(this.mnuGalleryApply);
            this.mnuGalleryClear.on('click', function () {
                me.api && me.api.ClearFormating();
                Common.NotificationCenter.trigger('edit:complete', toolbar);
            });
            this.mnuGalleryApply.on('click', function () { me.showApplyStyles(); });

            // item menu: Modify...
            if (toolbar.styleMenu) {
                this.mnuStyleModify = new Common.UI.MenuItem({caption: this.textModify});
                toolbar.styleMenu.addItem(this.mnuStyleModify);
                this.mnuStyleModify.on('click', function (item) {
                    var name = toolbar.styleMenu.items[0] && toolbar.styleMenu.items[0].styleTitle;
                    name && me.showModifyStyle(name);
                });
                toolbar.styleMenu.on('show:before', function () {
                    // shown for an item only (the first item is visible for an item)
                    me.mnuStyleModify.setVisible(toolbar.styleMenu.items[0].isVisible());
                });
            }
        },

        // ---------------------------------------------------------------- style lists
        getStyleList: function () {
            var api = this.api, me = this;
            if (!api || !api.asc_GetStylesArray) return [];
            var used = {};
            _.each(api.asc_getAllUsedParagraphStyles ? api.asc_getAllUsedParagraphStyles() : [], function (s) { used[s.get_Name()] = true; });
            var res = [];
            _.each(api.asc_GetStylesArray(), function (s) {
                var type = s.get_Type();
                if (type !== Asc.c_oAscStyleType.Paragraph && type !== Asc.c_oAscStyleType.Character) return;
                var name = s.get_Name();
                if (!s.get_QFormat() && !used[name]) return;
                var prio = s.get_UIPriority();
                (prio === undefined || prio === null || prio < 0) && (prio = (name === 'Normal') ? 0 : 99);
                res.push({
                    name: name,
                    caption: me.translate(name),
                    type: type === Asc.c_oAscStyleType.Character ? 'c' : 'p',
                    prio: prio,
                    isDefault: !!api.asc_IsStyleDefault(name),
                    isChanged: !!(api.asc_IsStyleDefault(name) && api.asc_IsDefaultStyleChanged(name))
                });
            });
            res.sort(function (a, b) {
                return a.prio !== b.prio ? a.prio - b.prio : a.caption.localeCompare(b.caption, undefined, {numeric: true, sensitivity: 'base'});
            });
            return res;
        },

        previewCss: function (name) {
            var info = this.api.asc_GetStyleForModify ? this.api.asc_GetStyleForModify(name) : null,
                tp = info && info.TextPr;
            if (!tp) return '';
            var css = [];
            tp.FontName && css.push('font-family: &quot;' + Common.Utils.String.htmlEncode(tp.FontName) + '&quot;');
            tp.FontSize && css.push('font-size: ' + Math.min(Math.max(tp.FontSize, 8), 16) + 'pt');
            tp.Bold && css.push('font-weight: bold');
            tp.Italic && css.push('font-style: italic');
            tp.Underline && css.push('text-decoration: underline');
            tp.SmallCaps && css.push('font-variant: small-caps');
            tp.Caps && css.push('text-transform: uppercase');
            var c = tp.Color;
            c && c.get_auto && !c.get_auto() && css.push('color: #' + Common.Utils.ThemeColor.getHexColor(c.get_r(), c.get_g(), c.get_b()));
            return css.join('; ');
        },

        // ---------------------------------------------------------------- Styles pane
        toggleStylesPane: function (show) {
            if (this.stylesPane && this.stylesPane.isVisible()) {
                if (show !== true) this.stylesPane.close();
                return;
            }
            var me = this;
            this.stylesPane = new DE.Views.StylesPane({
                handler: function (action, record) { me.onStylesPaneAction(action, record); }
            });
            this.stylesPane.on('close', function () { me.stylesPane = null; });
            this.stylesPane.show();
            this.refreshStylesPane();
        },

        refreshStylesPane: function () {
            if (!this.stylesPane) return;
            var me = this, list = this.getStyleList(), preview = this.stylesPane.isPreview();
            _.each(list, function (s) { s.preview = preview ? me.previewCss(s.name) : ''; });
            this.stylesPane.setStyles(list, this._currentStyle);
        },

        onStylesPaneAction: function (action, record) {
            var api = this.api, me = this, name = record ? record.get('name') : null;
            if (!api) return;
            switch (action) {
                case 'apply':
                    if (record.get('isClear')) {
                        api.ClearFormating();
                        api.put_Style('Normal');
                    } else
                        api.put_Style(name);
                    Common.NotificationCenter.trigger('edit:complete', this.toolbar);
                    break;
                case 'refresh':
                    this.refreshStylesPane();
                    break;
                case 'clear':
                    api.ClearFormating();
                    Common.NotificationCenter.trigger('edit:complete', this.toolbar);
                    break;
                case 'new':
                    this.showNewStyle();
                    break;
                case 'update':
                    var style = api.asc_GetStyleFromFormatting();
                    style.put_Name(name);
                    api.asc_AddNewStyle(style);
                    break;
                case 'modify':
                    this.showModifyStyle(name);
                    break;
                case 'delete':
                    Common.UI.warning({
                        msg: Common.Utils.String.format(this.textDeleteConfirm, record.get('caption')),
                        buttons: ['yes', 'no'],
                        primary: 'yes',
                        callback: function (btn) { btn === 'yes' && api.asc_RemoveStyle(name); }
                    });
                    break;
                case 'revert':
                    api.asc_RemoveStyle(name);
                    break;
                case 'pane':
                    this.toggleStylesPane(true);
                    break;
            }
        },

        // ---------------------------------------------------------------- Apply Styles box
        showApplyStyles: function () {
            if (this.applyStylesDlg && this.applyStylesDlg.isVisible()) {
                this.applyStylesDlg.setCurrent(this._currentStyle);
                return;
            }
            var me = this;
            this.applyStylesDlg = new DE.Views.ApplyStylesDialog({
                styles: this.getStyleList(),
                handler: function (action, name) {
                    if (action === 'apply') {
                        me.api.put_Style(name);
                    } else if (action === 'modify') {
                        me.showModifyStyle(name);
                    } else if (action === 'pane') {
                        me.toggleStylesPane(true);
                    }
                }
            });
            this.applyStylesDlg.on('close', function () { me.applyStylesDlg = null; });
            this.applyStylesDlg.show();
            this.applyStylesDlg.setCurrent(this._currentStyle);
        },

        // ---------------------------------------------------------------- Modify / New Style
        getFontStore: function () {
            if (!this.fontStore && this.toolbar && this.toolbar.cmbFontName) {
                this.fontStore = new Common.Collections.Fonts();
                this.fontStore.add(_.filter(_.map(this.toolbar.cmbFontName.store.toJSON(), _.clone), function (f) { return !f.cloneid; }));
            }
            return this.fontStore;
        },

        showModifyStyle: function (name) {
            var api = this.api, me = this;
            var info = api && api.asc_GetStyleForModify ? api.asc_GetStyleForModify(name) : null;
            if (!info) return;
            info.Caption = this.translate(info.Name);
            this.openStyleDialog(info, false, function (changes) {
                !_.isEmpty(changes) && api.asc_ModifyStyle(name, changes);
            });
        },

        showNewStyle: function () {
            var api = this.api, me = this;
            if (!api) return;
            var base = this._currentStyle || 'Normal',
                info = {
                    Name: '', Caption: this.findFreeStyleName(), Type: Asc.c_oAscStyleType.Paragraph, BasedOn: base, Next: null,
                    IsDefault: false, QFormat: true,
                    TextPr: api.asc_GetFontDialogProps ? api.asc_GetFontDialogProps() : {},
                    ParaPr: {}
                };
            this.openStyleDialog(info, true, function (changes, dlgName) {
                var title = dlgName || info.Caption,
                    style = api.asc_GetStyleFromFormatting();
                style.put_Name(title);
                style.get_Link() && style.get_Link().put_Name(title + '_character');
                style.put_Next(changes.Next || null);
                api.asc_AddNewStyle(style);
                delete changes.Name;
                delete changes.Next;
                !_.isEmpty(changes) && api.asc_ModifyStyle(title, changes);
            });
        },

        findFreeStyleName: function () {
            var names = {};
            _.each(this.getStyleList(), function (s) { names[s.name.toLowerCase()] = true; });
            for (var i = 1; i < 1000; i++)
                if (!names[(this.textStyleN + i).toLowerCase()]) return this.textStyleN + i;
            return this.textStyleN;
        },

        openStyleDialog: function (info, isNew, onOk) {
            var me = this, list = this.getStyleList();
            var dlg = new DE.Views.ModifyStyleDialog({
                isNew: isNew,
                styleInfo: info,
                styleNames: _.filter(list, function (s) { return s.type === 'p'; }),
                fontStore: this.getFontStore(),
                handler: function (result, changes) {
                    if (result === 'font') {
                        me.openFontDialogForStyle(dlg, changes);
                        return;
                    }
                    if (result === 'ok') {
                        var name = this.inputName ? this.inputName.getValue().trim() : '';
                        onOk(changes || {}, name);
                    }
                    Common.NotificationCenter.trigger('edit:complete', me.toolbar);
                }
            });
            dlg.show();
        },

        // Format > Font... of the Modify Style dialog (the Font dialog comes with the Font group)
        openFontDialogForStyle: function (dlg, textPr) {
            if (!DE.Views.FontDialog) return;
            var me = this;
            (new DE.Views.FontDialog({
                props: textPr,
                fontStore: this.getFontStore(),
                api: this.api,
                styleMode: true,
                handler: function (result, props) {
                    result === 'ok' && props && dlg.applyTextProps(props);
                }
            })).show();
        },

        textCreateStyle: 'Create a Style',
        textClearFormatting: 'Clear Formatting',
        textApplyStyles: 'Apply Styles...',
        textModify: 'Modify...',
        textDeleteConfirm: 'Do you want to delete the style {0} from the document?',
        textStyleN: 'Style'
    }, DE.Controllers.WordFormatting || {}));
});

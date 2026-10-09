/*
 * SPDX-FileCopyrightText: 2026 ElixirOffice contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 *
 * PE.Controllers.SlideSections: sections of a presentation, as in PowerPoint (sdkjs slide/Editor/sections.js).
 * Home > Slides > Section: Add Section (then the name box opens), Rename Section, Remove Section (Remove Section,
 * Remove All Sections, Remove Section & Slides), Collapse All, Expand All. The headers in the thumbnails and in the
 * Slide Sorter open a menu with the right mouse button. A section is the group of slides that follow its header;
 * the first slide of the presentation always starts the first section.
 */
define([
    'core'
], function () {
    'use strict';

    PE.Controllers.SlideSections = Backbone.Controller.extend({
        initialize: function () {
            Common.NotificationCenter.on('app:ready', this.onAppReady.bind(this));
        },

        onAppReady: function (mode) {
            if (this.ready || !mode || !mode.isEdit || mode.isEditDiagram || mode.isEditMailMerge || mode.isEditOle) return;
            this.ready = true;
            this.mode = mode;
            var me = this;
            this.api = this.getApplication().getController('Viewport').getApi();
            if (!this.api || !this.api.asc_GetSections) return;
            window.PE.SlideSections = this;
            this.addRibbonButton();
            this.api.asc_registerCallback('asc_onSectionContextMenu', _.bind(this.onContextMenu, this));
            this.api.asc_registerCallback('asc_onSectionsChanged', function () { me.updateMenu(); Common.NotificationCenter.trigger('slidesections:changed'); });
            this.api.asc_registerCallback('asc_onFocusObject', _.bind(this.updateMenu, this));
            this.api.asc_registerCallback('asc_onCountPages', _.bind(this.updateMenu, this));
        },

        // ---- Home > Slides > Section
        addRibbonButton: function () {
            var me = this, tb = this.getApplication().getController('Toolbar').toolbar;
            var $group = $('#slot-elx-reset-slide').closest('.group');
            if (!$group.length) return;
            $group.addClass('elx-three-rows').append($('<div class="elset"><span class="btn-slot text" id="slot-elx-section"></span></div>'));
            this.menu = this.buildMenu();
            this.btn = new Common.UI.Button({
                id: 'id-toolbar-btn-elx-section', cls: 'btn-toolbar', iconCls: 'toolbar__icon btn-elx-sb-sorter',
                caption: this.txtSection, hint: this.tipSection, menu: this.menu,
                lock: [Common.enumLock.lostConnect, Common.enumLock.noSlides, Common.enumLock.disableOnStart],
                dataHint: '1', dataHintDirection: 'bottom', dataHintOffset: 'small'
            });
            this.btn.render($('#slot-elx-section'));
            tb.lockControls && tb.lockControls.push(this.btn);
            this.menu.on('show:before', _.bind(this.updateMenu, this));
        },

        // the menu of the ribbon button
        buildMenu: function () {
            var me = this;
            var click = function (menu, item) { me.run(item.value); };
            return new Common.UI.Menu({
                items: [
                    {caption: this.txtAdd, value: 'add'},
                    {caption: this.txtRename, value: 'rename'},
                    {
                        caption: this.txtRemove, value: 'remove-sub',
                        menu: new Common.UI.Menu({
                            menuAlign: 'tl-tr',
                            items: [
                                {caption: this.txtRemove, value: 'remove'},
                                {caption: this.txtRemoveAll, value: 'remove-all'},
                                {caption: this.txtRemoveSlides, value: 'remove-slides'}
                            ]
                        }).on('item:click', click)
                    },
                    {caption: '--'},
                    {caption: this.txtCollapseAll, value: 'collapse-all'},
                    {caption: this.txtExpandAll, value: 'expand-all'}
                ]
            }).on('item:click', click);
        },

        // the section the menu acts on: the one given (header menu) or the one of the current slide
        sectionAt: function (id) {
            var sections = this.api.asc_GetSections() || [], cur = this.api.getCurrentPage();
            if (id !== undefined && id !== null) return _.find(sections, function (s) { return s.id === id; }) || null;
            return _.find(sections, function (s) { return cur >= s.first && cur < s.first + s.count; }) || null;
        },

        updateMenu: function () {
            if (!this.menu || !this.api) return;
            var sections = this.api.asc_GetSections() || [], has = sections.length > 0, here = this.sectionAt();
            _.each(this.menu.items, function (it) {
                switch (it.value) {
                    case 'rename': case 'remove-sub': it.setDisabled(!here); break;
                    case 'collapse-all': case 'expand-all': it.setDisabled(!has); break;
                }
            });
            var rm = _.find(this.menu.items, function (it) { return it.value === 'remove-sub'; });
            if (rm && rm.menu) _.each(rm.menu.items, function (it) {
                it.setDisabled(it.value === 'remove-all' ? !has : !here);
            });
        },

        run: function (cmd, id) {
            var api = this.api, s = this.sectionAt(id);
            switch (cmd) {
                case 'add':
                    var newId = api.asc_AddSection();
                    newId && this.rename(newId);
                    break;
                case 'rename': s && this.rename(s.id); break;
                case 'remove': s && api.asc_RemoveSection(s.id); break;
                case 'remove-all': api.asc_RemoveAllSections(); break;
                case 'remove-slides': s && api.asc_RemoveSectionAndSlides(s.id); break;
                case 'collapse-all': case 'expand-all':
                    _.each(api.asc_GetSections() || [], function (x) { api.asc_SetSectionCollapsed(x.id, cmd === 'collapse-all'); });
                    break;
                case 'up': s && api.asc_MoveSection(s.id, -1); break;
                case 'down': s && api.asc_MoveSection(s.id, 1); break;
            }
            Common.NotificationCenter.trigger('edit:complete');
        },

        rename: function (id) {
            var me = this, s = this.sectionAt(id);
            if (!s) return;
            new Common.Views.TextInputDialog({
                title: this.txtRenameTitle,
                label: this.txtSectionName,
                value: s.name || '',
                inputConfig: {
                    allowBlank: false,
                    validation: function (value) { return value.length < 255 ? true : me.txtLongName; }
                },
                handler: function (result, value) {
                    if (result === 'ok' && value) me.api.asc_RenameSection(id, value);
                    Common.NotificationCenter.trigger('edit:complete');
                }
            }).show();
        },

        // ---- the menu of a section header (thumbnails and Slide Sorter): {id, x, y} in the page
        onContextMenu: function (info) {
            if (!info || !this.mode.isEdit) return;
            this.showHeaderMenu(info.id, info.x, info.y);
        },

        showHeaderMenu: function (id, x, y) {
            var me = this, sections = this.api.asc_GetSections() || [];
            var index = _.findIndex(sections, function (s) { return s.id === id; });
            if (index < 0) return;
            this.hideHeaderMenu();
            var menu = new Common.UI.Menu({
                style: 'position: fixed; z-index: 10000;',
                items: [
                    {caption: this.txtRename, value: 'rename'},
                    {caption: this.txtRemove, value: 'remove'},
                    {caption: this.txtRemoveAll, value: 'remove-all'},
                    {caption: this.txtRemoveSlides, value: 'remove-slides'},
                    {caption: '--'},
                    {caption: this.txtMoveUp, value: 'up', disabled: index === 0},
                    {caption: this.txtMoveDown, value: 'down', disabled: index === sections.length - 1},
                    {caption: '--'},
                    {caption: this.txtCollapseAll, value: 'collapse-all'},
                    {caption: this.txtExpandAll, value: 'expand-all'}
                ]
            });
            menu.on('item:click', function (m, item) { me.run(item.value, id); });
            menu.on('hide:after', function () { _.defer(function () { me.hideHeaderMenu(); }); });
            menu.render($(document.body));
            menu.cmpEl.addClass('elx-section-menu');
            this._headerMenu = menu;
            menu.show();
            menu.cmpEl.css({position: 'fixed', left: x, top: y, display: 'block'});
        },

        hideHeaderMenu: function () {
            var m = this._headerMenu;
            this._headerMenu = null;
            if (m) { m.cmpEl && m.cmpEl.remove(); }
        },

        txtSection: 'Section',
        tipSection: 'Organize your slides into sections.',
        txtAdd: 'Add Section',
        txtRename: 'Rename Section',
        txtRemove: 'Remove Section',
        txtRemoveAll: 'Remove All Sections',
        txtRemoveSlides: 'Remove Section & Slides',
        txtCollapseAll: 'Collapse All',
        txtExpandAll: 'Expand All',
        txtMoveUp: 'Move Section Up',
        txtMoveDown: 'Move Section Down',
        txtRenameTitle: 'Rename Section',
        txtSectionName: 'Section name',
        txtLongName: 'The name is too long.'
    });
});

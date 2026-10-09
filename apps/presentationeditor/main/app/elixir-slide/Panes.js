/*
 * SPDX-FileCopyrightText: 2026 ElixirOffice contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 *
 * PE.Controllers.SlidePanes: the Selection pane of PowerPoint (Alt+F10, Option+F10 on macOS; Home > Arrange > Selection Pane;
 * sdkjs slide/Editor/selection-pane.js). The objects of the current slide, the one in front first: click
 * selects (Ctrl adds), the eye shows or hides (one undo step), double click renames (one undo step), Show All,
 * Hide All, Bring Forward, Send Backward. Esc in the pane or its close button closes it.
 */
define([
    'core'
], function () {
    'use strict';

    PE.Controllers.SlidePanes = Backbone.Controller.extend({
        initialize: function () {
            this._open = false;
            Common.NotificationCenter.on('app:ready', this.onAppReady.bind(this));
        },

        onAppReady: function (mode) {
            if (this.ready || !mode || mode.isEditDiagram || mode.isEditMailMerge || mode.isEditOle) return;
            this.ready = true;
            this.mode = mode;
            var me = this;
            this.api = this.getApplication().getController('Viewport').getApi();
            if (!this.api || !this.api.asc_GetSelectionPaneObjects) return;
            var refresh = _.debounce(function () { me.refresh(); }, 120);
            this.api.asc_registerCallback('asc_onFocusObject', refresh);
            this.api.asc_registerCallback('asc_onCountPages', refresh);
            this.api.asc_registerCallback('asc_onDocumentContentReady', refresh);
            this.api.asc_registerCallback('asc_onUserActionEnd', refresh);
            this.api.asc_registerCallback('asc_onUndoRedo', refresh);
            document.addEventListener('keydown', _.bind(this.onKey, this), true);
            // Esc closes the pane while the last click was in it (the editor takes the keyboard focus back after a click)
            document.addEventListener('mousedown', function (e) { me._paneFocus = !!$(e.target).closest('#elx-task-pane').length; }, true);
            this.addArrangeItem();
        },

        // Home > Arrange > Selection Pane
        addArrangeItem: function () {
            var me = this, tb = this.getApplication().getController('Toolbar').toolbar,
                btn = tb && tb.btnShapeArrange;
            if (!btn || !btn.menu) return;
            this.menuItem = new Common.UI.MenuItem({caption: this.txtSelectionPane, value: 'selection-pane', iconCls: 'menu__icon btn-select-all'});
            this.menuItem.on('click', function () { me.toggle(); });
            btn.menu.addItem({caption: '--'});
            btn.menu.addItem(this.menuItem);
        },

        onKey: function (e) {
            var K = Common.UI.Keys;
            if (e.keyCode === K.F10 && e.altKey && !e.ctrlKey && !e.shiftKey && !e.metaKey) {
                if (Common.Utils.ModalWindow.isVisible()) return;
                e.preventDefault(); e.stopPropagation();
                this.toggle();
            } else if (e.keyCode === K.ESC && this._open && (this._paneFocus || $(e.target).closest('#elx-task-pane').length) && !$(e.target).is('input')) {
                e.preventDefault(); e.stopPropagation();
                this.close();
            }
        },

        toggle: function () {
            this._open ? this.close() : this.open();
        },

        open: function () {
            if (!$('#elx-task-pane-style').length) {
                $('<style id="elx-task-pane-style"></style>').text(
                    '#elx-task-pane{position:absolute;right:0;top:0;bottom:0;width:290px;z-index:500;display:flex;flex-direction:column;' +
                    'background:var(--background-toolbar,#fff);color:var(--text-normal,#333);border-left:1px solid var(--border-toolbar,#ccc);box-shadow:-2px 0 6px rgba(0,0,0,.08);font-size:11px;}' +
                    '#elx-task-pane .elx-tp-head{display:flex;align-items:center;padding:8px 10px;font-size:14px;border-bottom:1px solid var(--border-divider,#e0e0e0);}' +
                    '#elx-task-pane .elx-tp-head span{flex:1;}' +
                    '#elx-task-pane .elx-tp-close{border:0;background:none;font-size:16px;cursor:pointer;color:inherit;padding:0 4px;}' +
                    '#elx-task-pane .elx-tp-body{flex:1;overflow:auto;padding:8px 10px;}' +
                    '#elx-task-pane .elx-sp-tools{display:flex;gap:6px;margin-bottom:6px;flex-wrap:wrap;}' +
                    '#elx-task-pane .elx-sp-tools button{font-size:11px;padding:2px 6px;}' +
                    '#elx-task-pane .elx-sp-item{display:flex;align-items:center;padding:3px 4px;cursor:pointer;border:1px solid transparent;}' +
                    '#elx-task-pane .elx-sp-item:hover{background:var(--highlight-button-hover,#e0e0e0);}' +
                    '#elx-task-pane .elx-sp-item.selected{background:var(--highlight-button-pressed,#cbcbcb);}' +
                    '#elx-task-pane .elx-sp-item.hidden-obj .elx-sp-name{opacity:.55;}' +
                    '#elx-task-pane .elx-sp-name{flex:1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}' +
                    '#elx-task-pane .elx-sp-eye{border:0;background:none;cursor:pointer;color:inherit;font-size:12px;width:22px;}' +
                    '#elx-task-pane .elx-sp-empty{color:var(--text-secondary,#777);margin-top:8px;}'
                ).appendTo('head');
            }
            this.close(true);
            var me = this, $host = $('#editor_sdk').parent();
            var $pane = $('<div id="elx-task-pane" role="complementary" tabindex="-1"><div class="elx-tp-head"><span></span>' +
                '<button class="elx-tp-close" type="button">&#x2715;</button></div><div class="elx-tp-body"></div></div>');
            $pane.attr('aria-label', this.txtSelection).attr('data-pane', 'selection').find('.elx-tp-head span').text(this.txtSelection);
            $pane.find('.elx-tp-close').attr('title', this.txtClose).on('click', function () { me.close(); });
            $host.css('position') === 'static' && $host.css('position', 'relative');
            $host.append($pane);
            this._open = true;
            this.refresh();
        },

        close: function (silent) {
            $('#elx-task-pane').remove();
            this._open = false;
            this._paneFocus = false;
            if (!silent && this.api) {
                this.api.asc_enableKeyEvents(true);
                Common.NotificationCenter.trigger('edit:complete');
            }
        },

        refresh: function () {
            if (!this._open || !this.api) return;
            var $body = $('#elx-task-pane .elx-tp-body');
            if (!$body.length || $body.find('input').length) return;
            this.render($body);
        },

        render: function ($body) {
            var me = this, list = this.api.asc_GetSelectionPaneObjects() || [], edit = !!(this.mode && this.mode.isEdit);
            var $tools = $('<div class="elx-sp-tools"></div>');
            if (edit) {
                $('<button type="button" class="btn btn-text-default" data-cmd="show"></button>').text(this.txtShowAll).appendTo($tools);
                $('<button type="button" class="btn btn-text-default" data-cmd="hide"></button>').text(this.txtHideAll).appendTo($tools);
                $('<button type="button" class="btn btn-text-default" data-cmd="forward"></button>').text(this.txtForward).attr('title', this.txtForward).appendTo($tools);
                $('<button type="button" class="btn btn-text-default" data-cmd="backward"></button>').text(this.txtBackward).attr('title', this.txtBackward).appendTo($tools);
            }
            var $list = $('<div class="elx-sp-list" role="listbox"></div>');
            _.each(list, function (o) {
                var $it = $('<div class="elx-sp-item" role="option"><span class="elx-sp-name"></span><button type="button" class="elx-sp-eye"></button></div>');
                $it.attr('data-id', o.id).toggleClass('selected', !!o.selected).toggleClass('hidden-obj', !!o.hidden).attr('aria-selected', !!o.selected);
                $it.find('.elx-sp-name').text(o.name);
                $it.find('.elx-sp-eye').html(o.hidden ? '&#x2014;' : '&#x1F441;').attr('title', o.hidden ? me.txtShow : me.txtHide).attr('aria-pressed', !o.hidden).prop('disabled', !edit);
                $list.append($it);
            });
            $body.empty().append($tools, $list);
            !list.length && $body.append($('<div class="elx-sp-empty"></div>').text(this.txtNoObjects));

            $list.find('.elx-sp-item').on('click', function (e) {
                if ($(e.target).closest('.elx-sp-eye').length) return;
                me.api.asc_SelectPaneObject($(this).attr('data-id'), e.ctrlKey || e.metaKey);
                me.refresh();
            }).on('dblclick', function () {
                if (!edit) return;
                var id = $(this).attr('data-id'), $name = $(this).find('.elx-sp-name'), old = $name.text();
                var $input = $('<input type="text" class="form-control" style="width:100%;height:20px;font-size:11px;">').val(old);
                $name.empty().append($input);
                me.api.asc_enableKeyEvents(false);
                $input.focus().select();
                var done = function (save) {
                    var v = $input.val().trim();
                    $input.off();
                    save && v && v !== old && me.api.asc_RenamePaneObject(id, v);
                    me.api.asc_enableKeyEvents(true);
                    me.render($body);
                };
                $input.on('keydown', function (e) {
                    if (e.keyCode === Common.UI.Keys.RETURN) { e.preventDefault(); e.stopPropagation(); done(true); }
                    else if (e.keyCode === Common.UI.Keys.ESC) { e.preventDefault(); e.stopPropagation(); done(false); }
                }).on('blur', function () { done(true); });
            });
            $list.find('.elx-sp-eye').on('click', function () {
                var $it = $(this).closest('.elx-sp-item');
                me.api.asc_SetPaneObjectsHidden($it.attr('data-id'), !$it.hasClass('hidden-obj'));
                me.refresh();
            });
            $tools.find('button').on('click', function () {
                var cmd = $(this).attr('data-cmd');
                if (cmd === 'show' || cmd === 'hide') me.api.asc_SetPaneObjectsHidden(null, cmd === 'hide');
                else me.api.asc_MovePaneObject(cmd === 'forward');
                me.refresh();
            });
        },

        txtSelection: 'Selection',
        txtSelectionPane: 'Selection Pane...',
        txtClose: 'Close',
        txtShow: 'Show',
        txtHide: 'Hide',
        txtShowAll: 'Show All',
        txtHideAll: 'Hide All',
        txtForward: 'Bring Forward',
        txtBackward: 'Send Backward',
        txtNoObjects: 'There are no objects on this slide.'
    });
});

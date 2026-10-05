/*
 * SPDX-FileCopyrightText: 2026 ElixirOffice contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 *
 * SSE.Controllers.ExcelGrid: Excel-like chrome around the grid, added to the
 * stock views without rewriting them.
 *
 * - Formula bar: Cancel (X) and Enter (check) buttons before "fx", enabled
 *   while a cell is being edited, like Excel.
 * - Status bar: Average, Count and Sum by default (Minimum and Maximum can be
 *   added from the right-click menu, which is remembered), Normal / Page Break
 *   Preview view buttons, a zoom slider and the zoom percentage, like Excel.
 */

define([
    'core',
    'common/main/lib/component/Slider'
], function () {
    'use strict';

    SSE.Controllers.ExcelGrid = Backbone.Controller.extend(_.extend({

        initialize: function () {
            Common.NotificationCenter.on('app:ready', _.bind(function (mode) { _.defer(_.bind(this.onAppReady, this, mode)); }, this));
        },

        ctrl: function (name) { return this.getApplication().getController(name); },

        icon: function (name) {
            var I = window.ElixirChromeIcons;
            return I ? I.svg(name, 20) : '';
        },

        onAppReady: function (mode) {
            this.mode = mode || {};
            var main = this.ctrl('Main');
            this.api = main && main.api;
            if (!this.api || this.ready) return;
            this.ready = true;
            $('body').addClass('elx-excel');
            this.setupFormulaBar();
            this.setupStatusbar();
        },

        // ---- formula bar ------------------------------------------------------------------
        setupFormulaBar: function () {
            var me = this, $name = $('#cell-editing-box .ce-group-name');
            if (!$name.length || $name.find('.elx-ce-btns').length) return;
            var $btns = $('<span class="elx-ce-btns">' +
                '<span class="elx-ce-grip" aria-hidden="true"></span>' +
                '<button id="ce-btn-cancel" type="button" class="btn tiny btn-toolbar disabled" aria-label="' + Common.Utils.String.htmlEncode(this.tipCancel) + '">' + this.icon('dismiss') + '</button>' +
                '<button id="ce-btn-enter" type="button" class="btn tiny btn-toolbar disabled" aria-label="' + Common.Utils.String.htmlEncode(this.tipEnter) + '">' + this.icon('checkmark') + '</button>' +
                '</span>');
            $name.find('#ce-func-label').before($btns);
            this.$ceCancel = $btns.find('#ce-btn-cancel');
            this.$ceEnter = $btns.find('#ce-btn-enter');
            this.$ceCancel.tooltip({title: this.tipCancel, placement: 'cursor'});
            this.$ceEnter.tooltip({title: this.tipEnter, placement: 'cursor'});
            // mousedown: act before the cell editor loses the focus
            this.$ceCancel.on('mousedown', function (e) {
                e.preventDefault();
                if (!$(this).hasClass('disabled')) { me.api.asc_closeCellEditor(true); Common.NotificationCenter.trigger('edit:complete'); }
            });
            this.$ceEnter.on('mousedown', function (e) {
                e.preventDefault();
                if (!$(this).hasClass('disabled')) { me.api.asc_closeCellEditor(false); Common.NotificationCenter.trigger('edit:complete'); }
            });
            this.api.asc_registerCallback('asc_onEditCell', function (state) {
                var editing = state !== Asc.c_oAscCellEditorState.editEnd;
                me.$ceCancel.toggleClass('disabled', !editing);
                me.$ceEnter.toggleClass('disabled', !editing);
            });
        },

        // ---- status bar -------------------------------------------------------------------
        setupStatusbar: function () {
            var sbc = this.ctrl('Statusbar'), sb = sbc && sbc.statusbar, me = this;
            if (!sb || !sb.boxZoom) return;
            this.statusbar = sb;

            // Excel shows Average, Count and Sum; Minimum and Maximum on demand
            var saved = Common.localStorage.getItem('sse-elx-statusbar-math');
            var shown = saved ? saved.split(',') : ['average', 'count', 'sum'];
            if (sb.customizeStatusBarMenu) {
                _.each(sb.customizeStatusBarMenu.items, function (it) {
                    if (!it.value || it.value === 'status') return;
                    var on = shown.indexOf(it.value) >= 0;
                    if (it.checked !== on) {
                        it.setChecked(on, true);
                        sb.onCustomizeStatusBarClick(null, it, {stopPropagation: function () {}});
                    }
                });
                sb.customizeStatusBarMenu.on('item:click', function (menu, item) {
                    if (!item.value || item.value === 'status') return;
                    var list = _.filter(menu.items, function (it) { return it.value && it.value !== 'status' && it.checked; });
                    Common.localStorage.setItem('sse-elx-statusbar-math', _.pluck(list, 'value').join(','));
                });
            }

            // view buttons + zoom slider + percentage, like Excel
            var $box = sb.boxZoom;
            $box.addClass('elx-zoom-box');
            var $views = $('<span class="elx-sb-views">' +
                '<button type="button" class="btn small btn-toolbar elx-sb-normal" data-hint="0" data-hint-direction="top" data-hint-offset="small">' + this.icon('table') + '</button>' +
                '<button type="button" class="btn small btn-toolbar elx-sb-pagebreak" data-hint="0" data-hint-direction="top" data-hint-offset="small">' + this.icon('pages') + '</button>' +
                '</span>');
            $box.find('#status-btn-zoomdown').before($views);
            $views.find('.elx-sb-normal').tooltip({title: this.tipNormal, placement: 'cursor'}).on('click', function () { me.setViewMode(Asc.c_oAscESheetViewType.normal); });
            $views.find('.elx-sb-pagebreak').tooltip({title: this.tipPageBreak, placement: 'cursor'}).on('click', function () { me.setViewMode(Asc.c_oAscESheetViewType.pageBreakPreview); });
            this.$views = $views;
            this.syncViewButtons();
            this.api.asc_registerCallback('asc_updateSheetViewType', _.bind(this.syncViewButtons, this));
            this.api.asc_registerCallback('asc_onActiveSheetChanged', _.bind(this.syncViewButtons, this));

            var $slider = $('<div id="status-zoom-slider" class="status-zoom-slider"></div>');
            $box.find('#status-btn-zoomdown').after($slider);
            $box.find('.cnt-zoom').insertAfter($box.find('#status-btn-zoomup'));
            if (Common.UI.SingleSlider) {
                this.zoomSlider = new Common.UI.SingleSlider({el: $slider, width: 100, minValue: 0, maxValue: 100, step: 1, value: 50});
                this.zoomSlider.on('change', function (s, v) { me._sliding = true; me.api.asc_setZoom(me.sliderToZoom(v) / 100); });
                this.zoomSlider.on('changecomplete', function (s, v) {
                    me._sliding = false;
                    me.api.asc_setZoom(me.sliderToZoom(v) / 100);
                    Common.NotificationCenter.trigger('edit:complete', sb);
                });
            }
            var update = function (zf) {
                var pr = Math.round((zf || me.api.asc_getZoom()) * 100);
                sb.labelZoom.text(Common.Utils.String.format(me.textZoomPercent, pr));
                me.zoomSlider && !me._sliding && me.zoomSlider.setValue(me.zoomToSlider(pr));
            };
            this.api.asc_registerCallback('asc_onZoomChanged', function (zf) { _.defer(function () { update(zf); }); });
            update();
            sb.labelZoom.css('min-width', 40);
            sbc.zoomText = this.textZoomPercent;
        },

        // like Excel: 10%-100% on the left half of the slider, 100%-400% on the right half
        zoomToSlider: function (p) {
            return p <= 100 ? Math.max(0, (p - 10) / 90 * 50) : Math.min(100, 50 + (p - 100) / 300 * 50);
        },
        sliderToZoom: function (pos) {
            var z = Math.round(pos <= 50 ? 10 + pos / 50 * 90 : 100 + (pos - 50) / 50 * 300);
            return Math.abs(z - 100) < 4 ? 100 : z;
        },

        setViewMode: function (type) {
            this.api.asc_SetSheetViewType && this.api.asc_SetSheetViewType(type);
            this.syncViewButtons();
            Common.NotificationCenter.trigger('edit:complete');
        },

        syncViewButtons: function () {
            if (!this.$views || !this.api) return;
            var t = this.api.asc_GetSheetViewType ? this.api.asc_GetSheetViewType(this.api.asc_getActiveWorksheetIndex()) : Asc.c_oAscESheetViewType.normal;
            this.$views.find('.elx-sb-normal').toggleClass('active', t === Asc.c_oAscESheetViewType.normal);
            this.$views.find('.elx-sb-pagebreak').toggleClass('active', t === Asc.c_oAscESheetViewType.pageBreakPreview);
        },

        tipCancel: 'Cancel',
        tipEnter: 'Enter',
        tipNormal: 'Normal',
        tipPageBreak: 'Page Break Preview',
        textZoomPercent: '{0}%'
    }, SSE.Controllers.ExcelGrid || {}));
});

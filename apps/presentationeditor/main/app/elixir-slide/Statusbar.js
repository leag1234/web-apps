/*
 * SPDX-FileCopyrightText: 2026 ElixirOffice contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 *
 * PE.Controllers.SlideStatusbar: PowerPoint-like status bar, built on the
 * stock one (PE.Views.Statusbar) without changing its template:
 * "Slide X of Y" and language (stock) | Notes | Comments | view buttons
 * Normal, Reading View, Slide Show | zoom out, zoom slider, zoom in,
 * percentage (zoom menu), Fit slide to current window.
 */
define([
    'core',
    'common/main/lib/component/Button',
    'common/main/lib/component/Slider'
], function () {
    'use strict';

    var STRINGS = {
        txtNotes: 'Notes', tipNotes: 'Notes: enter speaker notes below the slide',
        txtComments: 'Comments', tipComments: 'Show the Comments pane',
        tipNormal: 'Normal', tipReadingView: 'Reading View', tipSlideShow: 'Slide Show (Shift+F5)',
        tipFit: 'Fit slide to current window', tipZoom: 'Zoom', textZoomPercent: '{0}%'
    };

    PE.Views.SlideStatusbar = Backbone.View.extend(_.extend({ initialize: function () {} }, STRINGS, PE.Views.SlideStatusbar || {}));

    PE.Controllers.SlideStatusbar = Backbone.Controller.extend({

        initialize: function () {
            Common.NotificationCenter.on('app:ready', this.onAppReady.bind(this));
        },

        onAppReady: function (mode) {
            var me = this;
            if (this.ready) return;
            this.ready = true;
            this.s = new PE.Views.SlideStatusbar();
            this.api = this.getApplication().getController('Viewport').getApi();
            this.mode = mode || {};
            var $sb = $('#statusbar .statusbar');
            if (!$sb.length || !this.api) return;
            $sb.addClass('elx-slide-statusbar');
            this.build($sb);
            this.api.asc_registerCallback('asc_onZoomChange', function (percent) { me.onZoom(percent); });
            this.api.asc_registerCallback('asc_onNotesShow', function (show) { me.btnNotes && me.btnNotes.toggle(!!show, true); });
            this.api.asc_registerCallback('asc_onEndDemonstration', function () { me.syncViews(); });
            // current zoom: the stock label already shows it ("Zoom 85%")
            var m = /(\d+)/.exec($('#statusbar #status-label-zoom').text() || '');
            this.onZoom(m ? parseInt(m[1]) : 100);
        },

        btn: function ($slot, icon, opts) {
            var b = new Common.UI.Button(_.extend({
                cls: 'btn-toolbar' + (opts.caption ? '' : ' small'),
                iconCls: 'toolbar__icon ' + icon,
                dataHint: '0', dataHintDirection: 'top', dataHintOffset: 'small'
            }, opts));
            b.render($slot);
            return b;
        },

        build: function ($sb) {
            var me = this, s = this.s, $fit = $sb.find('#btn-zoom-topage');
            var $group = $fit.parent();
            var slot = function (id) { return $('<span class="elx-sb-slot" id="' + id + '"></span>'); };
            var $views = $('<span class="elx-sb-views"></span>').append(
                slot('elx-sb-notes'), slot('elx-sb-comments'), $('<span class="separator short elx-sb-sep"></span>'),
                slot('elx-sb-normal'), slot('elx-sb-reading'), slot('elx-sb-show'), $('<span class="separator short elx-sb-sep"></span>'));
            $group.find('#btn-zoom-down').before($views);

            this.btnNotes = this.btn($views.find('#elx-sb-notes'), 'btn-elx-sb-notes', { caption: s.txtNotes, hint: s.tipNotes, enableToggle: true });
            this.btnNotes.toggle(!Common.localStorage.getBool('pe-hidden-notes', false), true);
            this.btnNotes.on('click', function (b) {
                me.api.asc_ShowNotes(b.pressed);
                Common.localStorage.setBool('pe-hidden-notes', !b.pressed);
                var vt = me.getApplication().getController('ViewTab');
                vt && vt.view && vt.view.chNotes && vt.view.chNotes.setValue(b.pressed, true);
                Common.NotificationCenter.trigger('edit:complete');
            });
            if (this.mode.canViewComments !== false && this.mode.canCoAuthoring !== false) {
                this.btnComments = this.btn($views.find('#elx-sb-comments'), 'btn-elx-sb-comments', { caption: s.txtComments, hint: s.tipComments, enableToggle: true });
                this.btnComments.on('click', function () {
                    var c = me.getApplication().getController('ElixirChrome');
                    c && c.toggleComments();
                });
                Common.NotificationCenter.on('layout:changed', function () { me.syncComments(); });
            }
            this.btnNormal = this.btn($views.find('#elx-sb-normal'), 'btn-elx-sb-normal', { hint: s.tipNormal, enableToggle: true, allowDepress: false, toggleGroup: 'elx-sb-views' });
            this.btnNormal.toggle(true, true);
            this.btnReading = this.btn($views.find('#elx-sb-reading'), 'btn-elx-sb-reading', { hint: s.tipReadingView, enableToggle: true, allowDepress: false, toggleGroup: 'elx-sb-views' });
            this.btnShow = this.btn($views.find('#elx-sb-show'), 'btn-elx-sb-show', { hint: s.tipSlideShow });
            this.btnNormal.on('click', function () { me.syncViews(); });
            this.btnReading.on('click', function () { me.startShow(false); });
            this.btnShow.on('click', function () { me.startShow(true); });

            // zoom: - [slider] + 85% [fit]
            var $slider = $('<span class="elx-sb-zoom-slider" id="elx-sb-zoom-slider"></span>');
            $group.find('#btn-zoom-down').after($slider);
            $group.find('.cnt-zoom').insertAfter($group.find('#btn-zoom-up'));
            $fit.insertAfter($group.find('.cnt-zoom')).attr('title', s.tipFit);
            $group.find('#btn-zoom-towidth').addClass('elx-sb-hidden');
            this.slider = new Common.UI.SingleSlider({ el: $slider, width: 100, minValue: 0, maxValue: 100, step: 1, value: 50 });
            this.slider.on('change', function (sl, v) { me._sliding = true; me.api.zoom(me.sliderToZoom(v)); });
            this.slider.on('changecomplete', function (sl, v) { me._sliding = false; me.api.zoom(me.sliderToZoom(v)); Common.NotificationCenter.trigger('edit:complete'); });
        },

        // PowerPoint-like scale: 10%-100% on the left half, 100%-400% on the right half
        zoomToSlider: function (p) { return p <= 100 ? Math.max(0, (p - 10) / 90 * 50) : Math.min(100, 50 + (p - 100) / 300 * 50); },
        sliderToZoom: function (v) {
            var z = Math.round(v <= 50 ? 10 + v / 50 * 90 : 100 + (v - 50) / 50 * 300);
            return Math.abs(z - 100) < 4 ? 100 : z;
        },

        onZoom: function (percent) {
            var $l = $('#statusbar #status-label-zoom');
            $l.text(Common.Utils.String.format(this.s.textZoomPercent, percent)).css('min-width', 36);
            this.slider && !this._sliding && this.slider.setValue(this.zoomToSlider(percent));
        },

        // Reading View: the slide show in the window; Slide Show: full screen, like PowerPoint
        startShow: function (full) {
            var api = this.api;
            Common.NotificationCenter.trigger('preview:start', api.getCurrentPage(), false, false, true);
            if (full) {
                var vp = this.getApplication().getController('Viewport'), pv = vp && vp.getView('DocumentPreview');
                pv && pv.toggleFullScreen && pv.toggleFullScreen();
            }
            this.btnReading.toggle(true, true);
        },

        syncViews: function () { this.btnNormal && this.btnNormal.toggle(true, true); },

        syncComments: function () {
            var lm = this.getApplication().getController('LeftMenu');
            var on = !!(lm && lm.leftMenu && lm.leftMenu.btnComments && lm.leftMenu.btnComments.pressed);
            this.btnComments && this.btnComments.toggle(on, true);
        }
    });
});

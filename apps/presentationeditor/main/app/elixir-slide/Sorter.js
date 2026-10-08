/*
 * SPDX-FileCopyrightText: 2026 ElixirOffice contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 *
 * PE.SlideSorter: the Slide Sorter view of PowerPoint. A grid of every slide of
 * the presentation over the editing area (View > Slide Sorter, status bar
 * button): click selects (Ctrl toggles, Shift extends), double-click or Enter
 * opens the slide in the Normal view, drag and drop reorders (Ctrl copies),
 * Delete removes, Esc or the Normal button returns. The slide images come from
 * the engine (api.asc_getSlideImage), the moves go through the engine's own
 * shiftSlides so they are one undo step and co-authoring aware.
 */
define(['core'], function () {
    'use strict';

    var TEXT = { slide: 'Slide', title: 'Slide Sorter' };
    var THUMB_W = 280;

    var S = window.PE.SlideSorter = {
        isOpen: false,
        sel: [],
        anchor: 0,
        $el: null,
        dropPos: -1,

        api: function () { return window.PE.getController('Main').api; },
        count: function () { return this.api().getCountPages(); },

        toggle: function (on) { (on === undefined ? !this.isOpen : on) ? this.open() : this.close(); },

        open: function () {
            if (this.isOpen) return;
            var api = this.api(), me = this;
            if (!api || !api.getCountPages()) return;
            this.isOpen = true;
            this.sel = this.currentSelection();
            this.anchor = this.sel[0] || 0;
            this.$el = $('<div id="elx-sorter" tabindex="0" role="listbox" aria-label="' + TEXT.title + '"></div>').appendTo($('#editor-container'));
            this.$el.on('click', function (e) { if (e.target === me.$el[0] || $(e.target).is('.elx-sorter-grid')) me.select([], 0); });
            this.$el.on('keydown', function (e) { me.onKey(e); });
            this.$grid = $('<div class="elx-sorter-grid"></div>').appendTo(this.$el);
            this.render();
            api.asc_enableKeyEvents(false);
            this.$el.focus();
            this.countHandler = function () { if (me.isOpen) me.scheduleRender(); };
            api.asc_registerCallback('asc_onCountPages', this.countHandler);
            Common.NotificationCenter.trigger('slidesorter:changed', true);
        },

        close: function (goTo) {
            if (!this.isOpen) return;
            var api = this.api();
            this.isOpen = false;
            api.asc_unregisterCallback && api.asc_unregisterCallback('asc_onCountPages', this.countHandler);
            this.$el.remove();
            this.$el = null;
            api.asc_enableKeyEvents(true);
            if (typeof goTo === 'number' && goTo >= 0 && goTo < api.getCountPages()) api.goToPage(goTo);
            Common.NotificationCenter.trigger('slidesorter:changed', false);
            Common.NotificationCenter.trigger('edit:complete');
        },

        currentSelection: function () {
            var api = this.api();
            try {
                var arr = api.WordControl.Thumbnails.GetSelectedArray();
                if (arr && arr.length) return arr.slice();
            } catch (e) { /* fall back to the current slide */ }
            return [api.getCurrentPage()];
        },

        scheduleRender: function () {
            var me = this;
            clearTimeout(this._t);
            this._t = setTimeout(function () { me.isOpen && me.render(); }, 150);
        },

        render: function () {
            var me = this, api = this.api(), n = api.getCountPages(), i;
            this.sel = _.filter(this.sel, function (k) { return k < n; });
            if (!this.sel.length && n) this.sel = [Math.min(this.anchor, n - 1)];
            this.$grid.empty();
            var first = 1;
            try { first = api.asc_getFirstSlideNumber ? api.asc_getFirstSlideNumber() : 1; } catch (e) { /* default */ }
            for (i = 0; i < n; i++) {
                (function (idx) {
                    var $card = $('<div class="elx-sorter-card" role="option"></div>').attr('data-idx', idx);
                    $('<span class="elx-sorter-num"></span>').text(idx + (first || 1)).appendTo($card);
                    var $img = $('<img class="elx-sorter-img" alt="" draggable="false">').appendTo($card);
                    $card.attr('aria-label', TEXT.slide + ' ' + (idx + 1));
                    $card.on('click', function (e) { me.onCardClick(e, idx); });
                    $card.on('dblclick', function () { me.close(idx); });
                    $card.on('mousedown', function (e) { me.onPointerDown(e, idx); });
                    me.$grid.append($card);
                    // draw the images in small batches so a long deck does not freeze the editor
                    setTimeout(function () {
                        if (!me.isOpen) return;
                        var url = api.asc_getSlideImage(idx, THUMB_W * 2);
                        if (url) $img.attr('src', url);
                    }, Math.floor(idx / 6) * 20);
                })(i);
            }
            this.paintSelection();
        },

        paintSelection: function () {
            var me = this;
            this.$grid.children('.elx-sorter-card').each(function () {
                var idx = +$(this).attr('data-idx'), on = me.sel.indexOf(idx) >= 0;
                $(this).toggleClass('selected', on).attr('aria-selected', on);
            });
        },

        select: function (arr, anchor) {
            this.sel = arr;
            if (anchor !== undefined) this.anchor = anchor;
            this.paintSelection();
            if (arr.length) {
                this.api().asc_selectSlides(arr);
                var $c = this.$grid.children('[data-idx="' + arr[arr.length - 1] + '"]');
                $c.length && $c[0].scrollIntoView({ block: 'nearest' });
            }
        },

        onCardClick: function (e, idx) {
            var cur;
            if (this._justDragged) { e.stopPropagation(); return; }
            if (e.shiftKey) {
                cur = _.range(Math.min(this.anchor, idx), Math.max(this.anchor, idx) + 1);
                this.select(cur);
            } else if (e.ctrlKey || e.metaKey) {
                cur = this.sel.indexOf(idx) >= 0 ? _.without(this.sel, idx) : this.sel.concat([idx]);
                this.select(cur, idx);
            } else
                this.select([idx], idx);
            this.$el.focus();
            e.stopPropagation();
        },

        onKey: function (e) {
            var n = this.count(), last = this.sel.length ? this.sel[this.sel.length - 1] : 0, next = null,
                cols = Math.max(1, Math.floor(this.$grid.width() / this.$grid.children().first().outerWidth(true)) || 1);
            switch (e.keyCode) {
                case 27: this.close(); break;                                  // Esc
                case 13: this.close(this.sel.length ? this.sel[0] : 0); break; // Enter
                case 37: next = last - 1; break;
                case 39: next = last + 1; break;
                case 38: next = last - cols; break;
                case 40: next = last + cols; break;
                case 36: next = 0; break;
                case 35: next = n - 1; break;
                case 46: case 8: this.remove(); break;
                case 65: if (e.ctrlKey || e.metaKey) this.select(_.range(n), 0); else return; break;
                default: return;
            }
            if (next !== null && next >= 0 && next < n) {
                if (e.shiftKey) this.select(_.range(Math.min(this.anchor, next), Math.max(this.anchor, next) + 1));
                else if ((e.ctrlKey || e.metaKey) && (e.keyCode === 37 || e.keyCode === 39)) this.moveSelection(e.keyCode === 37 ? -1 : 1);
                else this.select([next], next);
            }
            e.preventDefault();
            e.stopPropagation();
        },

        // Ctrl+Left / Ctrl+Right: move the selected slides one place
        moveSelection: function (dir) {
            var n = this.count(), s = this.sel.slice().sort(function (a, b) { return a - b; });
            if (dir < 0 && s[0] === 0 || dir > 0 && s[s.length - 1] === n - 1) return;
            this.api().asc_moveSlidesTo(s, dir < 0 ? s[0] - 1 : s[s.length - 1] + 2, false);
            this.sel = _.map(s, function (k) { return k + dir; });
            this.api().asc_selectSlides(this.sel);
            this.render();
        },

        remove: function () {
            var api = this.api();
            if (!this.sel.length || this.sel.length >= this.count()) return;
            api.asc_selectSlides(this.sel);
            api.DeleteSlide();
            this.sel = [Math.max(0, Math.min(this.sel[0], this.count() - 2))];
            this.scheduleRender();
        },

        // Drag and drop with pointer events (works the same in every browser): the gap of the
        // nearest card half under the pointer is the drop position, Ctrl / Alt copies.
        onPointerDown: function (e, idx) {
            if (e.button !== 0) return;
            var me = this, x0 = e.clientX, y0 = e.clientY, started = false;
            var move = function (ev) {
                if (!started) {
                    if (Math.abs(ev.clientX - x0) + Math.abs(ev.clientY - y0) < 6) return;
                    started = true;
                    if (me.sel.indexOf(idx) < 0) me.select([idx], idx);
                    me.dragging = true;
                    me.$grid.addClass('dragging');
                }
                me.updateDrop(ev);
            };
            var up = function (ev) {
                $(document).off('mousemove.elxsorter mouseup.elxsorter');
                if (!started) return;
                me.updateDrop(ev);
                me.finishDrop(!!(ev.ctrlKey || ev.altKey));
                // the click that follows the mouse up must not change the selection
                me._justDragged = true;
                setTimeout(function () { me._justDragged = false; }, 0);
            };
            $(document).on('mousemove.elxsorter', move).on('mouseup.elxsorter', up);
        },

        updateDrop: function (ev) {
            var best = null, bestD = 1e9, me = this;
            this.$grid.children('.elx-sorter-card').each(function () {
                var r = this.getBoundingClientRect(), cx = r.left + r.width / 2, cy = r.top + r.height / 2,
                    d = Math.abs(ev.clientY - cy) * 3 + Math.abs(ev.clientX - cx);
                if (d < bestD) { bestD = d; best = { el: this, idx: +$(this).attr('data-idx'), after: ev.clientX > cx }; }
            });
            this.$grid.children().removeClass('drop-before drop-after');
            if (!best) { this.dropPos = -1; return; }
            this.dropPos = best.idx + (best.after ? 1 : 0);
            $(best.el).addClass(best.after ? 'drop-after' : 'drop-before');
        },

        clearDrop: function () {
            this.dragging = false;
            this.dropPos = -1;
            if (this.$grid) this.$grid.removeClass('dragging').children().removeClass('drop-before drop-after');
        },

        finishDrop: function (copy) {
            var pos = this.dropPos;
            this.clearDrop();
            if (pos < 0 || !this.sel.length) return;
            var s = this.sel.slice().sort(function (a, b) { return a - b; });
            this.api().asc_moveSlidesTo(s, pos, copy);
            // the moved slides end up together at the drop position
            var before = _.filter(s, function (k) { return k < pos; }).length;
            var start = copy ? pos : pos - before;
            this.sel = _.range(start, start + s.length);
            this.api().asc_selectSlides(this.sel);
            this.render();
        }
    };

    return S;
});

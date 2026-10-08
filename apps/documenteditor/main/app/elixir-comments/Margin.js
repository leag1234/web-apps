/*
 * SPDX-FileCopyrightText: 2026 ElixirOffice contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 *
 * ElixirOffice comment balloons in the right margin (document editor), like Word:
 * every comment is shown next to the text it is anchored to, in a markup area kept
 * on the right of the pages (sdkjs asc_ElixirSetMarkupWidth). The selected comment
 * gets a connector line to its text and is aligned with it, the others move aside.
 * Reply, resolve, edit and delete inline. The Comments pane stays available: while it
 * is open, the balloons hide (Word's "List" view), and the upstream's floating
 * popover is used only for a new comment that is still being typed.
 *
 * Data: Common.Controllers.Comments (collection of Common.Models.Comment); anchors:
 * api.asc_ElixirGetCommentAnchors(); refresh: asc_onElixirCommentAnchorsChanged
 * (scroll, zoom, recalculation) and the collection events.
 */
define([
    'core',
    'common/main/lib/controller/Comments'
], function () {
    'use strict';

    var WIDTH = 296,        // markup area, CSS px (narrower down to MIN_WIDTH when the window is narrow)
        MIN_WIDTH = 216,
        GAP = 8,            // vertical gap between two balloons
        PAD = 14;           // space between the page edge and the balloons

    var esc = function (s) { return Common.Utils.String.htmlEncode(s || '').replace(/\n/g, '<br>'); };
    var icon = function (name) { return window.ElixirChromeIcons ? window.ElixirChromeIcons.svg(name, 20) : ''; };

    DE.Views.ElixirCommentMargin = Backbone.View.extend({

        initialize: function () {
            this.on_ = false;
            this.active = null;     // uid of the selected comment
            this.hover = null;      // uid of the comment under the mouse in the text
            this.editing = null;    // { uid, mode: 'reply' | 'edit' }
            this.items = {};        // uid -> { $el, sig }
            this.drafts = {};       // uid -> text being typed
        },

        setup: function (api, mode, ctrl) {
            var me = this;
            this.api = api;
            this.mode = mode;
            this.ctrl = ctrl;
            this.canComment = !!mode.canComments && !mode.isLightVersion;

            var $main = $('#id_main');
            if (!$main.length || !api.asc_ElixirGetCommentAnchors) return;
            this.$layer = $('<div id="elx-markup" class="elx-markup" aria-label="' + this.textComments + '"></div>').appendTo($main);
            this.$svg = $('<svg class="elx-markup-lines" xmlns="http://www.w3.org/2000/svg"><path class="elx-markup-line"></path></svg>').appendTo(this.$layer);
            this.$balloons = $('<div class="elx-markup-balloons"></div>').appendTo(this.$layer);
            this.$menu = $('<div class="elx-cbal-menu" role="menu" style="display:none;"></div>').appendTo(this.$layer);

            var schedule = _.bind(this.schedule, this);
            api.asc_registerCallback('asc_onElixirCommentAnchorsChanged', schedule);
            api.asc_registerCallback('asc_onHideComment', function (hint) { me.onHideComment(hint); });
            ctrl.collection.on('add remove reset change', schedule);
            Common.NotificationCenter.on('layout:changed', schedule);
            Common.NotificationCenter.on('window:resize', schedule);
            $(window).on('resize', schedule);
            _.each(['asc_showComments', 'asc_hideComments'], function (fn) {
                var orig = api[fn];
                if (orig) api[fn] = function () { var r = orig.apply(api, arguments); schedule(); return r; };
            });

            this.$balloons.on('mousedown', '.elx-cbal', function (e) {
                // keep the document from taking the focus back
                if (!$(e.target).is('textarea')) e.preventDefault();
            });
            this.$balloons.on('click', '.elx-cbal', function (e) { me.onBalloonClick(e, $(this).attr('data-uid')); });
            this.$balloons.on('focus', 'textarea', function () { api.asc_enableKeyEvents(false); });
            this.$balloons.on('blur', 'textarea', function () { api.asc_enableKeyEvents(true); });
            this.$balloons.on('input', 'textarea', function () {
                var uid = $(this).closest('.elx-cbal').attr('data-uid');
                me.drafts[uid] = this.value;
                me.autosize(this);
                $(this).closest('.elx-cbal').find('.elx-cbal-post').prop('disabled', !this.value.trim());
            });
            this.$balloons.on('keydown', 'textarea', function (e) {
                var uid = $(this).closest('.elx-cbal').attr('data-uid');
                if (e.keyCode === 13 && (e.ctrlKey || e.metaKey)) { e.preventDefault(); me.submit(uid); }
                else if (e.keyCode === 27) { e.preventDefault(); me.cancelEdit(uid); }
                e.stopPropagation();
            });
            this.$menu.on('mousedown', function (e) { e.preventDefault(); });
            this.$menu.on('click', '[data-action]', function (e) {
                e.stopPropagation();
                var uid = me.$menu.attr('data-uid'), action = $(this).attr('data-action');
                me.hideMenu();
                if (action === 'edit') me.startEdit(uid, 'edit');
                else if (action === 'delete') me.ctrl.onRemoveComment(uid);
                else if (action === 'resolve') me.ctrl.onResolveComment(uid);
            });
            $(document).on('mousedown.elxmarkup', function (e) { if (!$(e.target).closest('.elx-cbal-menu, .elx-cbal-more').length) me.hideMenu(); });
            // the wheel over the balloons scrolls the document
            this.$layer.on('wheel', function (e) {
                var wc = api.WordControl, oe = e.originalEvent;
                if (wc && wc.m_oScrollVerApi && !$(e.target).closest('textarea').length) {
                    wc.m_oScrollVerApi.scrollBy(0, oe.deltaMode === 1 ? oe.deltaY * 40 : oe.deltaY);
                    e.preventDefault();
                }
            });
            this.schedule();
        },

        isOn: function () { return this.on_; },

        schedule: function () {
            var me = this;
            if (this.pending) return;
            this.pending = true;
            (window.requestAnimationFrame || setTimeout)(function () { me.pending = false; me.render(); });
        },

        // ---- state from the editor ---------------------------------------------------------------------------
        showResolved: function () { return !!Common.Utils.InternalSettings.get('de-settings-resolvedcomment'); },
        liveComments: function () { var v = Common.Utils.InternalSettings.get('de-settings-livecomment'); return v === undefined || v === null ? true : !!v; },
        paneOpen: function () { return $('#left-panel-comments').is(':visible'); },

        visibleItems: function () {
            var me = this, resolved = this.showResolved();
            var anchors = this.api.asc_ElixirGetCommentAnchors() || [];
            return _.filter(_.map(anchors, function (a) { return { a: a, c: me.ctrl.findComment(a.Id) }; }), function (it) {
                return it.c && !it.c.get('hide') && (resolved || !it.c.get('resolved'));
            });
        },

        wanted: function (items) {
            var wc = this.api.WordControl;
            if (!this.liveComments() || this.paneOpen() || !items.length) return false;
            if (wc && wc.isUseMultiPageView && wc.isUseMultiPageView()) return false;
            if (this.api.isViewMode && !this.mode.canComments && !this.mode.canViewComments) return false;
            return true;
        },

        // as wide as Word's balloons when the window has room for it, never narrower than MIN_WIDTH
        markupWidth: function () {
            var dd = this.api.WordControl && this.api.WordControl.m_oDrawingDocument, pg = dd && dd.m_arrPages && dd.m_arrPages[0];
            var pageW = pg && pg.drawingPage ? pg.drawingPage.right - pg.drawingPage.left : 816;
            var room = $('#id_main_view').width() - pageW - 40;
            return Math.max(MIN_WIDTH, Math.min(WIDTH, room));
        },

        // ---- layout -----------------------------------------------------------------------------------------
        render: function () {
            if (!this.api || !this.$layer) return;
            var items = this.visibleItems(), want = this.wanted(items), width = want ? this.markupWidth() : 0;
            if (want !== this.on_ || Math.abs(width - (this.width || 0)) > 4) {
                this.on_ = want;
                this.width = width;
                $('body').toggleClass('elx-markup-on', want);
                this.api.asc_ElixirSetMarkupWidth(width);
                if (!want) { this.active = null; this.editing = null; this.hideMenu(); }
                items = this.visibleItems();
            }
            this.$layer.toggle(this.on_);
            if (!this.on_) { this.$balloons.empty(); this.items = {}; return; }

            // clip to the document view (not over the rulers and the scroll bar)
            var $view = $('#id_main_view'), main = $('#id_main')[0].getBoundingClientRect(), view = $view[0].getBoundingClientRect();
            this.$layer.css({ left: view.left - main.left, top: view.top - main.top, width: view.width, height: view.height });
            var ox = view.left - main.left, oy = view.top - main.top;

            items.sort(function (x, y) { return (x.a.Page - y.a.Page) || (x.a.Y - y.a.Y) || (x.a.X - y.a.X); });
            var me = this, seen = {};
            _.each(items, function (it) {
                var uid = it.a.Id, rec = me.items[uid], sig = me.signature(it.c, uid);
                seen[uid] = true;
                if (!rec) {
                    rec = me.items[uid] = { $el: $('<div class="elx-cbal" role="article" tabindex="-1"></div>').attr('data-uid', uid) };
                    me.$balloons.append(rec.$el);
                }
                if (rec.sig !== sig) { me.fill(rec.$el, it.c, uid); rec.sig = sig; }
                it.$el = rec.$el;
                it.h = rec.$el.outerHeight();
                it.top = it.a.Y - oy - 6;
            });
            _.each(_.keys(this.items), function (uid) { if (!seen[uid]) { me.items[uid].$el.remove(); delete me.items[uid]; } });

            // stack the balloons; the selected one stays next to its text
            var k = _.findIndex(items, function (it) { return it.a.Id === me.active; }), i;
            if (k < 0) k = 0;
            if (items.length) {
                for (i = k + 1; i < items.length; i++) items[i].top = Math.max(items[i].top, items[i - 1].top + items[i - 1].h + GAP);
                for (i = k - 1; i >= 0; i--) items[i].top = Math.min(items[i].top, items[i + 1].top - GAP - items[i].h);
                if (me.active === null) // nothing selected: never push a balloon above its text
                    for (i = 1; i < items.length; i++) items[i].top = Math.max(items[i].top, items[i - 1].top + items[i - 1].h + GAP);
            }
            var line = '';
            _.each(items, function (it) {
                var left = it.a.PageRight - ox + PAD;
                it.$el.css({ top: Math.round(it.top), left: Math.round(left), width: me.width - 2 * PAD });
                var sel = it.a.Id === me.active, hov = it.a.Id === me.hover;
                it.$el.toggleClass('active', sel).toggleClass('hover', hov && !sel);
                if (sel || (hov && !line)) {
                    var y0 = it.a.Y - oy + it.a.H, x0 = it.a.X - ox, xe = it.a.PageRight - ox;
                    line = 'M' + x0 + ',' + y0 + 'H' + xe + 'L' + left + ',' + (it.top + 16);
                }
            });
            this.$svg.find('path').attr('d', line);
            this.$layer.toggleClass('elx-markup-hasactive', !!this.active);
        },

        signature: function (c, uid) {
            var r = c.get('replys') || [];
            return [c.get('comment'), c.get('username'), c.get('date'), c.get('resolved'), c.get('lock'), r.length,
                _.map(r, function (x) { return x.get('reply') + x.get('date'); }).join('|'),
                uid === this.active, this.editing && this.editing.uid === uid ? this.editing.mode : ''].join('\u0001');
        },

        person: function (m) {
            var name = AscCommon.UserInfoParser.getParsedName(m.get('username')) || m.get('username') || '',
                initials = m.get('initials') || Common.Utils.getUserInitials(name),
                color = m.get('usercolor') || '#7a7a7a';
            return '<span class="elx-cbal-av" style="background-color:' + color + ';">' + esc(initials) + '</span>' +
                '<span class="elx-cbal-who"><span class="elx-cbal-name">' + esc(name) + '</span>' +
                (m.get('date') ? '<span class="elx-cbal-date">' + esc(m.get('date')) + '</span>' : '') + '</span>';
        },

        fill: function ($el, c, uid) {
            var me = this, active = uid === this.active, ed = this.editing && this.editing.uid === uid ? this.editing.mode : null,
                replies = c.get('replys') || [], locked = !!c.get('lock'), html = [];
            $el.toggleClass('resolved', !!c.get('resolved'));
            html.push('<div class="elx-cbal-head">' + this.person(c));
            if (this.canComment && !locked) {
                html.push('<button type="button" class="elx-cbal-btn elx-cbal-resolve" data-act="resolve" title="' + (c.get('resolved') ? this.textReopen : this.textResolve) + '" aria-label="' + (c.get('resolved') ? this.textReopen : this.textResolve) + '">' + icon('checkmark') + '</button>');
                html.push('<button type="button" class="elx-cbal-btn elx-cbal-more" data-act="more" title="' + this.textMore + '" aria-label="' + this.textMore + '">' + icon('more') + '</button>');
            }
            html.push('</div>');
            if (ed === 'edit') {
                html.push('<div class="elx-cbal-box"><textarea class="elx-cbal-input" rows="2">' + Common.Utils.String.htmlEncode(this.drafts[uid] !== undefined ? this.drafts[uid] : c.get('comment')) + '</textarea>' +
                    '<div class="elx-cbal-actions"><button type="button" class="btn normal primary elx-cbal-post" data-act="save">' + this.textSave + '</button>' +
                    '<button type="button" class="btn normal elx-cbal-cancel" data-act="cancel">' + this.textCancel + '</button></div></div>');
            } else {
                html.push('<div class="elx-cbal-text' + (active ? '' : ' clamp') + '">' + esc(c.get('comment')) + '</div>');
            }
            if (replies.length) {
                if (active) {
                    _.each(replies, function (r) {
                        html.push('<div class="elx-cbal-reply"><div class="elx-cbal-head">' + me.person(r) + '</div><div class="elx-cbal-text">' + esc(r.get('reply')) + '</div></div>');
                    });
                } else {
                    html.push('<div class="elx-cbal-count">' + Common.Utils.String.format(replies.length > 1 ? this.textReplies : this.textReply1, replies.length) + '</div>');
                }
            }
            if (active && this.canComment && !locked && ed !== 'edit' && !c.get('resolved')) {
                var draft = this.drafts[uid] || '';
                html.push('<div class="elx-cbal-box"><textarea class="elx-cbal-input" rows="1" placeholder="' + this.textReplyPh + '" aria-label="' + this.textReplyPh + '">' + Common.Utils.String.htmlEncode(draft) + '</textarea>' +
                    '<div class="elx-cbal-actions"' + (ed === 'reply' || draft ? '' : ' style="display:none;"') + '><button type="button" class="btn normal primary elx-cbal-post" data-act="reply"' + (draft.trim() ? '' : ' disabled') + '>' + this.textPost + '</button>' +
                    '<button type="button" class="btn normal elx-cbal-cancel" data-act="cancel">' + this.textCancel + '</button></div></div>');
            }
            $el.html(html.join(''));
            $el.find('textarea').each(function () { me.autosize(this); });
            if (ed) _.defer(function () { var t = $el.find('textarea')[0]; if (t) { t.focus(); t.setSelectionRange(t.value.length, t.value.length); } });
        },

        autosize: function (t) { t.style.height = 'auto'; t.style.height = Math.min(t.scrollHeight + 2, 240) + 'px'; },

        // ---- interaction -----------------------------------------------------------------------------------
        onBalloonClick: function (e, uid) {
            var $t = $(e.target), act = $t.closest('[data-act]').attr('data-act');
            if (act === 'resolve') { this.ctrl.onResolveComment(uid); return; }
            if (act === 'more') { this.showMenu(uid, $t.closest('[data-act]')); return; }
            if (act === 'reply') { this.submit(uid); return; }
            if (act === 'save') { this.submit(uid); return; }
            if (act === 'cancel') { this.cancelEdit(uid); return; }
            if ($t.is('textarea')) {
                if (!this.editing || this.editing.uid !== uid) { this.editing = { uid: uid, mode: 'reply' }; }
                $t.closest('.elx-cbal-box').find('.elx-cbal-actions').show();
                return;
            }
            if (uid !== this.active) this.select(uid);
        },

        select: function (uid) {
            this.active = uid;
            this.editing = null;
            var c = this.ctrl.findComment(uid);
            if (c && this.api) {
                this.api.asc_selectComment(uid);
                this.ctrl._dontScrollToComment = true;
                this.api.asc_showComment(uid, false);
            }
            this.schedule();
        },

        startEdit: function (uid, mode) {
            this.active = uid;
            this.editing = { uid: uid, mode: mode };
            delete this.drafts[uid];
            this.schedule();
        },

        cancelEdit: function (uid) {
            delete this.drafts[uid];
            this.editing = null;
            if (this.items[uid]) this.items[uid].sig = null;
            this.api.asc_enableKeyEvents(true);
            this.schedule();
        },

        submit: function (uid) {
            var $el = this.items[uid] && this.items[uid].$el, text = $el ? ($el.find('textarea').val() || '').trim() : '';
            if (!text) return;
            var mode = this.editing && this.editing.uid === uid ? this.editing.mode : 'reply';
            delete this.drafts[uid];
            this.editing = null;
            if (this.items[uid]) this.items[uid].sig = null;
            if (mode === 'edit') this.ctrl.onChangeComment(uid, text);
            else this.ctrl.onAddReplyComment(uid, text);
            this.api.asc_enableKeyEvents(true);
            this.schedule();
        },

        showMenu: function (uid, $btn) {
            var c = this.ctrl.findComment(uid);
            if (!c) return;
            var items = [];
            if (c.get('editable')) items.push('<a role="menuitem" data-action="edit">' + this.textEdit + '</a>');
            items.push('<a role="menuitem" data-action="resolve">' + (c.get('resolved') ? this.textReopen : this.textResolve) + '</a>');
            if (c.get('removable')) items.push('<a role="menuitem" data-action="delete">' + this.textDelete + '</a>');
            var lay = this.$layer[0].getBoundingClientRect(), b = $btn[0].getBoundingClientRect();
            this.$menu.html(items.join('')).attr('data-uid', uid).css({ top: b.bottom - lay.top + 2, left: Math.max(0, b.right - lay.left - 160) }).show();
        },

        hideMenu: function () { this.$menu && this.$menu.hide(); },

        // ---- events from the editor -------------------------------------------------------------------------
        // the cursor or the mouse is on commented text (patched Comments controller)
        onShowComment: function (uids, hint) {
            var me = this, uid = _.find(uids, function (id) { return !!me.items[id]; }) || uids[0];
            if (hint) { if (this.hover !== uid) { this.hover = uid; this.schedule(); } return; }
            if (this.active !== uid) { this.active = uid; if (!this.editing || this.editing.uid !== uid) this.editing = null; this.schedule(); }
        },

        onHideComment: function (hint) {
            if (!this.on_) return;
            if (hint) { if (this.hover) { this.hover = null; this.schedule(); } return; }
            var $f = $(document.activeElement);
            if ($f.closest('.elx-cbal').length) return;   // typing a reply
            if (this.active || this.hover) { this.active = null; this.hover = null; this.editing = null; this.schedule(); }
        },

        textComments: 'Comments',
        textResolve: 'Resolve thread',
        textReopen: 'Reopen thread',
        textMore: 'More thread actions',
        textEdit: 'Edit comment',
        textDelete: 'Delete thread',
        textReplyPh: 'Reply',
        textPost: 'Post',
        textSave: 'Save',
        textCancel: 'Cancel',
        textReply1: '1 reply',
        textReplies: '{0} replies'
    });

    // The cursor or the mouse on commented text: with the balloons shown, select the balloon instead of
    // opening the floating popover. The prototype is patched before the controller binds its API callbacks.
    var proto = Common.Controllers.Comments.prototype,
        origShow = proto.onApiShowComment,
        origUpdate = proto.onApiUpdateCommentPosition;
    var margin = function () { return DE.Views.ElixirCommentMargin.instance; };
    proto.onApiShowComment = function (uids, posX, posY, leftX, opts, hint) {
        var m = margin();
        if (m && m.isOn() && !this.isDummyComment && window.DE && this.appPrefix === 'de-') {
            this.uids = _.clone(uids);
            this.isSelectedComment = !hint;
            this.popover && this.popover.isVisible() && this.popover.hide();
            m.onShowComment(uids, hint);
            return;
        }
        return origShow.apply(this, arguments);
    };
    proto.onApiUpdateCommentPosition = function () {
        var m = margin();
        if (m && m.isOn() && !this.isDummyComment && this.appPrefix === 'de-') return;
        return origUpdate.apply(this, arguments);
    };

    return DE.Views.ElixirCommentMargin;
});

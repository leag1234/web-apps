/*
 * ElixirOffice window chrome: Word-like Navigation pane and side panels.
 *
 * - The left icon rail goes away, like in Word: the side panels open on demand
 *   (Ctrl+F, View > Navigation, the Comments button of the title bar) and the
 *   left column takes no room when they are closed.
 * - Navigation pane with Headings / Pages / Results tabs on top of the upstream
 *   headings tree and search panel; Pages draws page thumbnails with the print
 *   preview renderer. Ctrl+F opens it on Results, like Word.
 * - Ctrl+wheel and pinch zoom smoothly (10 % per wheel notch) instead of
 *   jumping between preset steps.
 * - The comments pane docks on the right of the page, like Word's; the page
 *   area shrinks and stays centred.
 * - The properties panel on the right starts closed (View > Right panel).
 *
 * Licensed under the GNU AGPL v3 like the rest of web-apps.
 */
define([
    'core',
    'common/main/lib/elixir-chrome/Icons'
], function () {
    'use strict';

    var Icons = window.ElixirChromeIcons;
    var esc = function (s) { return Common.Utils.String.htmlEncode(s === undefined || s === null ? '' : String(s)); };

    var STRINGS = {
        txtNavigation: 'Navigation',
        txtClose: 'Close',
        txtHeadings: 'Headings',
        txtPages: 'Pages',
        txtResults: 'Results',
        txtPage: 'Page {0}'
    };

    DE.Views.ElixirNavPane = Backbone.View.extend(_.extend({

        initialize: function (options) {
            this.controller = options.controller;
            this.tab = 'headings';
        },

        icon: function (name, size, cls) { return Icons ? Icons.svg(name, size, cls) : ''; },

        // called before the document loads: no rail, right panel closed by default
        prepare: function () {
            $('body').addClass('elx-norail');
            try {
                if (Common.localStorage.getItem('de-hidden-rightmenu') === null)
                    Common.localStorage.setBool('de-hidden-rightmenu', true);
            } catch (e) {}
            this.bindKeys();
        },

        setup: function (mode) {
            var main = DE.getController('Main');
            this.api = main.api;
            this.mode = mode;
            this.$left = $('#left-menu');
            this.$panel = this.$left.find('.left-panel');
            if (!this.$panel.length || this.ready) return;
            this.ready = true;
            this.renderHead();
            this.dockComments();
            this.watchLeftWidth();
            this.watchPanels();
            this.bindZoom();
            this.sync();
            var me = this;
            this.api.asc_registerCallback('asc_onCountPages', function (n) { me.pageCount = n; me.tab === 'pages' && me.isOpen() && me.renderPages(); });
            this.api.asc_registerCallback('asc_onCurrentPage', function (n) { me.markCurrent(n); });
            this.api.asc_registerCallback('asc_onZoomChange', function (percent) { me.zoomValue = percent; });
        },

        lm: function () { var c = DE.getController('LeftMenu'); return c && c.leftMenu; },
        lmc: function () { return DE.getController('LeftMenu'); },

        // ---- the left column takes no room while its panels are closed ---------------------

        watchLeftWidth: function () {
            var me = this, el = this.$left[0];
            var check = this.checkLeft = function () {
                var w = parseInt(el.style.width) || 0;
                // closed, or open for a panel that lives elsewhere (the docked comments)
                var collapsed = w <= 40 || !me.$panel.children('.content-box').filter(function () { return this.style.display !== 'none'; }).length;
                if (collapsed !== me.$left.hasClass('elx-left-collapsed')) {
                    me.$left.toggleClass('elx-left-collapsed', collapsed);
                    _.defer(function () { Common.NotificationCenter.trigger('layout:changed', 'leftmenu'); });
                }
            };
            new MutationObserver(check).observe(el, { attributes: true, attributeFilter: ['style'] });
            check();
        },

        // ---- Navigation pane ---------------------------------------------------------------

        renderHead: function () {
            var me = this;
            this.$head = $('<div class="elx-navhead">' +
                '<div class="elx-navhead-top"><span class="elx-navhead-title">' + esc(this.txtNavigation) + '</span>' +
                    '<button type="button" class="elx-navhead-close" title="' + esc(this.txtClose) + '" aria-label="' + esc(this.txtClose) + '">' + this.icon('dismiss', 20) + '</button></div>' +
                '<div class="elx-navtabs" role="tablist">' +
                    '<button type="button" role="tab" data-tab="headings">' + esc(this.txtHeadings) + '</button>' +
                    '<button type="button" role="tab" data-tab="pages">' + esc(this.txtPages) + '</button>' +
                    '<button type="button" role="tab" data-tab="results">' + esc(this.txtResults) + '</button>' +
                '</div></div>');
            this.$panel.prepend(this.$head);
            this.$pages = $('<div id="left-panel-elx-pages" class="content-box elx-pages" style="display:none;"><div class="elx-pages-list"></div></div>');
            this.$panel.append(this.$pages);
            this.$head.find('[data-tab]').on('click', function () { me.openTab($(this).attr('data-tab')); });
            this.$head.find('.elx-navhead-close').on('click', function () { me.close(); });
            this.$pages.on('click', '.elx-page-thumb', function () {
                me.api.goToPage(parseInt($(this).attr('data-page')));
            });
        },

        watchPanels: function () {
            var me = this, timer;
            new MutationObserver(function () {
                clearTimeout(timer);
                timer = setTimeout(function () { me.sync(); me.checkLeft && me.checkLeft(); }, 0);
            }).observe(this.$panel[0], { attributes: true, attributeFilter: ['style', 'class'], subtree: true, childList: false });
        },

        // ---- comments pane on the right, like Word -----------------------------------------

        dockComments: function () {
            var me = this, $c = $('#left-panel-comments');
            if (!$c.length) return;
            this.$dock = $('<div id="elx-comments-dock" class="elx-comments-dock side-panel" style="display:none;"></div>').appendTo('#editor-container');
            this.$dock.append($c);
            var sync = function () {
                var open = $c[0].style.display !== 'none';
                if (open === me.$dock.is(':visible')) return;
                me.$dock.toggle(open);
                $('#editor-container').toggleClass('elx-comments-docked', open);
                me.checkLeft && me.checkLeft();
                _.defer(function () { Common.NotificationCenter.trigger('layout:changed', 'leftmenu'); });
            };
            new MutationObserver(sync).observe($c[0], { attributes: true, attributeFilter: ['style'] });
            sync();
        },

        isOpen: function () { return $('body').hasClass('elx-nav-on'); },

        sync: function () {
            var lm = this.lm();
            if (!lm) return;
            var nav = lm.btnNavigation && lm.btnNavigation.pressed,
                search = lm.btnSearchBar && lm.btnSearchBar.pressed;
            if (search) this.tab = 'results';
            else if (nav && this.tab === 'results') this.tab = 'headings';
            var on = !!(nav || search);
            $('body').toggleClass('elx-nav-on', on);
            this.$panel.toggleClass('elx-pages-mode', on && nav && this.tab === 'pages');
            this.$pages.toggle(on && nav && this.tab === 'pages');
            var current = on ? (search ? 'results' : this.tab) : '';
            this.$head.find('[data-tab]').each(function () {
                var active = $(this).attr('data-tab') === current;
                $(this).toggleClass('active', active).attr('aria-selected', active);
            });
            if (on && nav && this.tab === 'pages' && !this._pagesRendered) this.renderPages();
        },

        open: function (tab) { this.openTab(tab || 'headings'); },

        openTab: function (tab, findText) {
            var lmc = this.lmc(), lm = this.lm();
            if (!lmc || !lm) return;
            this.tab = tab;
            if (tab === 'results') {
                lmc.onShowHideSearch(true, findText);
            } else {
                if (!lm.btnNavigation.pressed) lmc.onShowHideNavigation(true);
                this._pagesRendered = false;
            }
            this.sync();
        },

        close: function () {
            var lm = this.lm();
            if (!lm) return;
            lm.close();
            Common.NotificationCenter.trigger('layout:changed', 'leftmenu');
            Common.NotificationCenter.trigger('edit:complete');
        },

        bindKeys: function () {
            var me = this;
            // Ctrl+F (Cmd+F) opens the Navigation pane on its Results tab, like Word
            window.addEventListener('keydown', function (e) {
                if (!(e.keyCode === 70 || e.code === 'KeyF')) return;
                var mod = Common.Utils.isMac ? e.metaKey && !e.ctrlKey : e.ctrlKey && !e.metaKey;
                if (!mod || e.altKey || e.shiftKey || !me.ready) return;
                var fm = me.lm() && me.lm().menuFile;
                if (fm && fm.isVisible()) return;
                if ($(e.target).closest('.elx-search, .modal-dlg, .asc-window').length) return;
                e.preventDefault();
                e.stopImmediatePropagation();
                Common.UI.Menu.Manager.hideAll();
                var sel = me.api && me.api.asc_GetSelectedText ? $.trim(me.api.asc_GetSelectedText() || '') : '';
                me.openTab('results', sel && sel.length < 200 ? sel : undefined);
                var lm = me.lm();
                setTimeout(function () { lm && lm.panelSearch && lm.panelSearch.focus && lm.panelSearch.focus('search'); }, 50);
            }, true);
        },

        // ---- Pages: thumbnails drawn with the print preview renderer ------------------------

        renderPages: function () {
            var api = this.api, me = this;
            if (!api) return;
            var n = this.pageCount || (api.getCountPages ? api.getCountPages() : 0) || 0, cur = api.getCurrentPage ? api.getCurrentPage() : 0;
            var $list = this.$pages.find('.elx-pages-list');
            var html = '';
            for (var i = 0; i < n; i++)
                html += '<button type="button" class="elx-page-thumb' + (i === cur ? ' current' : '') + '" data-page="' + i + '" aria-label="' + esc(Common.Utils.String.format(this.txtPage, i + 1)) + '">' +
                    '<canvas width="1" height="1"></canvas><span>' + (i + 1) + '</span></button>';
            $list.html(html);
            this._pagesRendered = true;
            clearTimeout(this._pagesTimer);
            var i0 = 0;
            var step = function () {
                // the File menu print preview owns the renderer while it is open
                if (api.printPreview || !me.$pages.is(':visible')) return;
                var end = Math.min(n, i0 + 4);
                for (; i0 < end; i0++) me.drawThumb(i0, $list.find('.elx-page-thumb[data-page=' + i0 + '] canvas')[0]);
                if (i0 < n) me._pagesTimer = setTimeout(step, 16);
            };
            this._pagesTimer = setTimeout(step, 30);
        },

        drawThumb: function (i, canvas) {
            var api = this.api;
            if (!canvas || !api.asc_initPrintPreview) return;
            var size = api.asc_getPageSize ? api.asc_getPageSize(i) : null;
            var wmm = size && size['W'] || 210, hmm = size && size['H'] || 297;
            var W = 150, H = Math.round(W * hmm / wmm), off = 25;
            var $host = $('#elx-thumb-render');
            if (!$host.length) $host = $('<div id="elx-thumb-render" aria-hidden="true"></div>').appendTo(document.body);
            $host.css({ width: W + 2 * off, height: H + 2 * off });
            try {
                api.asc_initPrintPreview('elx-thumb-render');
                api.asc_drawPrintPreview(i);
                var src = $host.find('canvas')[0];
                if (src) {
                    var r = src.width / (W + 2 * off);
                    canvas.width = Math.round(W * r);
                    canvas.height = Math.round(H * r);
                    canvas.getContext('2d').drawImage(src, Math.round(off * r), Math.round(off * r), canvas.width, canvas.height, 0, 0, canvas.width, canvas.height);
                }
            } catch (e) {
            } finally {
                api.asc_closePrintPreview && api.asc_closePrintPreview();
            }
        },

        markCurrent: function (n) {
            if (!this.$pages) return;
            this.$pages.find('.elx-page-thumb').removeClass('current').filter('[data-page=' + n + ']').addClass('current');
        },

        // ---- smooth zoom ---------------------------------------------------------------------

        bindZoom: function () {
            var me = this, target = null, raf = 0;
            var host = document.getElementById('editor_sdk');
            if (!host) return;
            host.addEventListener('wheel', function (e) {
                if (!e.ctrlKey || e.metaKey || !me.api || !me.api.zoom) return;
                e.preventDefault();
                e.stopPropagation();
                var unit = e.deltaMode === 1 ? 33 : e.deltaMode === 2 ? 400 : 1;
                // 100 px of wheel (one notch) = 10 %, trackpad pinches give small continuous deltas
                var factor = Math.exp(-e.deltaY * unit * Math.log(1.1) / 100);
                var base = target !== null ? target : (me.zoomValue || 100);
                target = Math.max(10, Math.min(500, base * factor));
                if (!raf) raf = requestAnimationFrame(function () {
                    raf = 0;
                    var z = Math.round(target);
                    if (Math.abs(z - (me.zoomValue || 0)) >= 1) me.api.zoom(z);
                    clearTimeout(me._zoomIdle);
                    me._zoomIdle = setTimeout(function () { target = null; }, 250);
                });
            }, { capture: true, passive: false });
        }

    }, STRINGS, DE.Views.ElixirNavPane || {}));
});

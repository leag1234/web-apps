/*
 * ElixirOffice: "Show Quick Access Toolbar" (shared by the Word, Excel, PowerPoint and PDF editors).
 *
 * Like Word's "Hide Quick Access Toolbar": when the toolbar is hidden, the quick access buttons
 * (and the AutoSave switch next to them) disappear from the title row and the row shrinks to the
 * default title height, which gives the document more vertical space. The file name, the search
 * box and the account / window controls stay.
 *
 * - the right-click menu on the quick access toolbar or on the ribbon tab row offers
 *   "Hide Quick Access Toolbar" / "Show Quick Access Toolbar";
 * - the choice is kept in the editor's local storage; a host (the desktop app) can set it with
 *   ElixirQuickAccess.set(visible, { host: true }) and listens to the "elixir:quickaccess" event
 *   on the editor window to save a change made here.
 *
 * Licensed under the GNU AGPL v3 like the rest of web-apps.
 */
define([
    'core',
    'notification'
], function () {
    'use strict';

    var KEY = 'elx-show-quick-access';
    var COMPACT_HEIGHT = 28;        // title row height without the quick access toolbar
    var visible = true;
    var originalHeight = 0;
    var $menu = null;

    var CSS =
        'body.elx-no-qat #box-document-title > .hedset[role="menubar"],' +
        'body.elx-no-qat #box-document-title .elx-autosave { display: none !important; }' +
        // the 28px row: controls are 24px high so that nothing touches the row edges
        'body.elx-no-qat .elx-search .elx-search-input { height: 24px !important; }' +
        'body.elx-no-qat .elx-search .elx-search-ic { top: 2px !important; }' +
        'body.elx-no-qat .elx-docname { height: 24px !important; line-height: 24px !important; }' +
        'body.elx-no-qat #box-document-title .color-user-name { width: 24px !important; height: 24px !important; min-width: 24px !important; line-height: 24px !important; }' +
        '.elx-qa-menu { position: fixed; z-index: 20000; min-width: 220px; padding: 4px; border-radius: 4px;' +
        ' background: var(--background-normal, #fff); color: var(--text-normal, #444);' +
        ' border: 1px solid var(--border-toolbar, #cbcbcb); box-shadow: 0 4px 16px rgba(0,0,0,.2); }' +
        '.elx-qa-menu button { display: block; width: 100%; padding: 6px 16px; border: 0; border-radius: 3px; background: transparent;' +
        ' color: inherit; font: inherit; font-size: 12px; text-align: start; white-space: nowrap; cursor: pointer; }' +
        '.elx-qa-menu button:hover, .elx-qa-menu button:focus-visible { background: var(--highlight-button-hover, #e0e0e0); outline: none; }';

    function app() { return window.DE || window.SSE || window.PE || window.PDFE || window.VE || null; }

    function viewport() {
        var a = app(), c = a && a.getController && a.getController('Viewport');
        return c || null;
    }

    function titleItem() {
        var c = viewport(), l = c && c.viewport && c.viewport.vlayout;
        return l && l.getItem ? { layout: l, item: l.getItem('title') } : null;
    }

    function readStored() {
        try { return Common.localStorage.getBool(KEY, true); } catch (e) { return true; }
    }

    function store(v) {
        try { Common.localStorage.setBool(KEY, v); } catch (e) {}
    }

    // the title row is as high as the default title height without the quick access toolbar
    function applyLayout() {
        var t = titleItem();
        if (!t || !t.item || !t.item.el || !t.item.el.is(':visible') || $('body').hasClass('eo-compact-editor')) return;
        var current = t.item.height;
        if (!current) return;
        if (!originalHeight) originalHeight = current;
        var target = visible ? originalHeight : Math.min(originalHeight, COMPACT_HEIGHT);
        if (current === target) return;
        t.item.height = target;
        Common.Utils.InternalSettings.set('document-title-height', target);
        t.layout.doLayout();
        Common.NotificationCenter.trigger('layout:changed', 'header');
        Common.NotificationCenter.trigger('window:resize');
    }

    function apply() {
        $('body').toggleClass('elx-no-qat', !visible);
        applyLayout();
    }

    function closeMenu() {
        if (!$menu) return;
        $menu.remove();
        $menu = null;
        $(document).off('.elxqa');
        $(window).off('.elxqa');
    }

    function label() {
        var c = viewport(), h = c && c.header;
        return (visible ? (h && h.txtHideQuickAccess) || 'Hide Quick Access Toolbar'
                        : (h && h.txtShowQuickAccess) || 'Show Quick Access Toolbar');
    }

    function openMenu(x, y) {
        closeMenu();
        Common.UI.Menu && Common.UI.Menu.Manager && Common.UI.Menu.Manager.hideAll();
        $menu = $('<div class="elx-qa-menu" role="menu"><button type="button" role="menuitem"></button></div>').appendTo(document.body);
        $menu.find('button').text(label()).on('click', function () {
            closeMenu();
            api.set(!visible);
            Common.NotificationCenter.trigger('edit:complete');
        });
        var w = $menu.outerWidth(), h = $menu.outerHeight();
        $menu.css({ left: Math.max(4, Math.min(x, window.innerWidth - w - 4)), top: Math.max(4, Math.min(y, window.innerHeight - h - 4)) });
        $(document).on('mousedown.elxqa', function (e) { if (!$(e.target).closest('.elx-qa-menu').length) closeMenu(); });
        $(document).on('keydown.elxqa', function (e) { if (e.keyCode === 27) { e.preventDefault(); closeMenu(); Common.NotificationCenter.trigger('edit:complete'); } });
        $(window).on('blur.elxqa resize.elxqa', closeMenu);
        setTimeout(function () { $menu && $menu.find('button').focus(); }, 0);
    }

    // right click on the quick access toolbar or on the empty part and the tabs of the ribbon tab row
    function onContextMenu(e) {
        if ($('body').hasClass('eo-compact-editor')) return;
        var $t = $(e.target);
        if ($t.closest('.elx-qa-menu').length) { e.preventDefault(); return; }
        var onBar = !!$t.closest('#box-document-title > .hedset[role="menubar"]').length;
        var onTabs = false;
        var $tb = $('#toolbar');
        if (!onBar && $tb.length && $t.closest('#toolbar').length &&
            !$t.closest('.btn, button, input, select, textarea, .dropdown-menu, .combobox, .form-control').length) {
            var top = $tb.offset().top - $(window).scrollTop(),
                tabs = Common.Utils.InternalSettings.get('toolbar-height-tabs') || 32;
            onTabs = e.clientY >= top && e.clientY < top + tabs;
        }
        if (!onBar && !onTabs) return;
        e.preventDefault();
        e.stopPropagation();
        openMenu(e.clientX, e.clientY);
    }

    var api = {
        isVisible: function () { return visible; },

        // options.host: the change comes from the host application (it is not reported back)
        set: function (value, options) {
            value = value !== false;
            var changed = value !== visible;
            visible = value;
            store(value);
            apply();
            if (changed && !(options && options.host)) {
                try { window.dispatchEvent(new CustomEvent('elixir:quickaccess', { detail: { visible: visible } })); } catch (e) {}
            }
        },

        // re-applies the state, for example after the compact desktop title bar gives the title row back
        refresh: function () { apply(); }
    };
    window.ElixirQuickAccess = api;
    Common.UI = Common.UI || {};
    Common.UI.ElixirQuickAccess = api;

    var style = document.createElement('style');
    style.textContent = CSS;
    document.head.appendChild(style);
    visible = readStored();
    document.addEventListener('contextmenu', onContextMenu, true);
    Common.NotificationCenter.on('app:ready', function () {
        var t = titleItem();
        if (t && t.item && t.item.height) originalHeight = t.item.height;
        apply();
    });
    // the class is set at once so that the quick access toolbar never flashes
    $('body').toggleClass('elx-no-qat', !visible);

    return api;
});

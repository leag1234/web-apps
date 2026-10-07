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
 *  LanguageDialog.js
 *
 *  Created on 04/25/2017
 *
 */
if (Common === undefined)
    var Common = {};

define([], function () { 'use strict';

    // Word-like Language dialog: the language of the cursor (or selection) on top, the recently used
    // languages, a search field, then the whole list (native name, English name below, a check on
    // the current one). The footer has "Set As Default" (document default language) and, below the
    // list, "Do not check spelling" (the editor switch for spell checking, see Review > Spelling).
    // handler(result, value, extra): result 'ok' | 'default' | 'cancel', value = language name
    // (e.g. "en-US"), extra = {code, noSpell}.
    Common.Views.LanguageDialog = Common.UI.Window.extend(_.extend({

    options: {
        header: true,
        width: 400,
        cls: 'modal-dlg lang-dlg',
        id: 'window-language'
    },

    initialize : function(options) {
        var opts = _.extend({}, this.options, options || {});
        var buttons = [];
        opts.showDefault && buttons.push({value: 'default', caption: this.txtSetDefault, cls: 'lang-dlg-default'});
        buttons.push('ok', 'cancel');
        opts.buttons = buttons;
        opts.title = opts.title || this.txtTitle;
        opts.footerCls = 'center';
        opts.tpl = _.template([
            '<div class="box">',
                '<label class="input-label lang-dlg-label"><%= label %></label>',
                '<div id="lang-dlg-search"></div>',
                '<div class="lang-dlg-list" id="lang-dlg-list" role="listbox" tabindex="-1"></div>',
                '<% if (showNoSpell) { %><div id="lang-dlg-nospell"></div><% } %>',
            '</div>'
        ].join(''))({label: this.labelMark, showNoSpell: opts.showNoSpell !== false});
        this.options = opts;
        Common.UI.Window.prototype.initialize.call(this, opts);
    },

    render: function() {
        Common.UI.Window.prototype.render.call(this);

        var me = this, $window = this.getChild();
        $window.find('.dlg-btn').on('click', _.bind(this.onBtnClick, this));

        this.languages = this.options.languages || [];
        this.current = _.findWhere(this.languages, {code: parseInt(this.options.current)});
        this.selected = this.current || this.languages[0];
        this.selKey = this.selected ? (this.current ? 'current:' : 'all:') + this.selected.code : null;
        this.$list = $window.find('#lang-dlg-list');

        this.inputSearch = new Common.UI.InputField({
            el: $window.find('#lang-dlg-search'),
            allowBlank: true,
            placeHolder: this.txtSearch,
            validateOnChange: true,
            style: 'width: 100%;',
            value: ''
        });
        this.inputSearch.on('changing', function(input, value) { me.filter = $.trim(value || '').toLowerCase(); me.fillList(); });
        this.inputSearch.cmpEl && this.inputSearch.cmpEl.find('input').attr({'aria-label': this.txtSearch, role: 'searchbox'}).on('keydown', _.bind(this.onSearchKey, this));

        this.chNoSpell = null;
        if (this.options.showNoSpell !== false) {
            this.chNoSpell = new Common.UI.CheckBox({
                el: $window.find('#lang-dlg-nospell'),
                labelText: this.txtNoSpell
            });
            this.chNoSpell.setValue(!!this.options.noSpell, true);
        }

        this.$list.on('click', '.lang-item', function(e) { me.select($(e.currentTarget).attr('data-key')); });
        this.$list.on('dblclick', '.lang-item', function(e) { me.select($(e.currentTarget).attr('data-key')); me.finish('ok'); });
        this.fillList();
        setTimeout(function() { me.inputSearch && me.inputSearch.focus(); }, 100);
        return this;
    },

    getRecent: function() {
        var key = 'app-settings-recent-langs',
            count = Common.Utils.InternalSettings.get(key + '-count') || 5,
            arr = Common.localStorage.getItem(key);
        arr = arr ? arr.split(';') : [];
        var langs = this.languages;
        return _.compact(_.map(arr.slice(0, count), function(v) { return _.findWhere(langs, {value: v}); }));
    },

    rememberRecent: function(value) {
        var key = 'app-settings-recent-langs',
            count = Common.Utils.InternalSettings.get(key + '-count') || 5,
            arr = Common.localStorage.getItem(key);
        arr = arr ? arr.split(';') : [];
        arr = _.without(arr, value);
        arr.unshift(value);
        arr.splice(count);
        Common.localStorage.setItem(key, arr.join(';'));
    },

    itemHtml: function(rec, section) {
        var isCur = this.current && rec.code === this.current.code,
            key = section + ':' + rec.code,
            isSel = key === this.selKey,
            en = (rec.displayValueEn && rec.displayValueEn !== rec.displayValue) ? rec.displayValueEn : '';
        return '<div class="lang-item' + (isSel ? ' selected' : '') + (isCur ? ' current' : '') + '" role="option" aria-selected="' + (isSel ? 'true' : 'false') + '" data-key="' + key + '">' +
                '<svg class="lang-check" viewBox="0 0 16 16" aria-hidden="true"' + (isCur ? '' : ' style="visibility:hidden"') + '><polyline points="3,8.5 6.5,12 13,4.5" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"/></svg>' +
                '<span class="lang-name">' + _.escape(rec.displayValue) + (en ? '<span class="lang-en">' + _.escape(en) + '</span>' : '') + '</span>' +
                (rec.spellcheck ? '<svg class="icon uni-scale spellcheck-lang" aria-hidden="true"><use href="#btn-ic-docspell"></use></svg>' : '') +
            '</div>';
    },

    fillList: function() {
        var me = this, html = '', f = this.filter,
            head = function(t) { return '<div class="lang-head" role="presentation">' + _.escape(t) + '</div>'; };
        if (f) {
            var found = _.filter(this.languages, function(r) { return (r.displayValue || '').toLowerCase().indexOf(f) > -1 || (r.displayValueEn || '').toLowerCase().indexOf(f) > -1 || (r.value || '').toLowerCase().indexOf(f) > -1; });
            if (found.length && !_.some(found, function(r) { return 'found:' + r.code === me.selKey; })) {
                this.selKey = 'found:' + found[0].code;
                this.selected = found[0];
            }
            html = found.length ? _.map(found, function(r) { return me.itemHtml(r, 'found'); }).join('') : '<div class="lang-empty">' + _.escape(this.txtNoResult) + '</div>';
        } else {
            if (this.selected && this.selKey.indexOf('found:') === 0)
                this.selKey = (this.current && this.current.code === this.selected.code ? 'current:' : 'all:') + this.selected.code;
            if (this.current) html += head(this.txtCurrent) + this.itemHtml(this.current, 'current');
            var recent = this.getRecent();
            if (recent.length) html += head(this.txtRecent) + _.map(recent, function(r) { return me.itemHtml(r, 'recent'); }).join('');
            html += head(this.txtAll) + _.map(this.languages, function(r) { return me.itemHtml(r, 'all'); }).join('');
        }
        this.$list.html(html);
        this.$list.scrollTop(0);
    },

    visibleKeys: function() {
        return this.$list.find('.lang-item').map(function() { return $(this).attr('data-key'); }).get();
    },

    select: function(key) {
        var rec = _.findWhere(this.languages, {code: parseInt(String(key).split(':')[1])});
        if (!rec) return;
        this.selected = rec;
        this.selKey = key;
        this.$list.find('.lang-item').removeClass('selected').attr('aria-selected', 'false')
            .filter(function() { return $(this).attr('data-key') === key; }).addClass('selected').attr('aria-selected', 'true');
        this.scrollToSelected();
    },

    scrollToSelected: function() {
        var $s = this.$list.find('.lang-item.selected').first();
        if (!$s.length) return;
        var $l = this.$list, top = $s.position().top + $l.scrollTop(), h = $l.height(), ih = $s.outerHeight();
        if (top < $l.scrollTop() + 28) $l.scrollTop(Math.max(0, top - 28));
        else if (top + ih > $l.scrollTop() + h) $l.scrollTop(top + ih - h);
    },

    onSearchKey: function(e) {
        var d = e.keyCode === Common.UI.Keys.DOWN ? 1 : (e.keyCode === Common.UI.Keys.UP ? -1 : 0);
        if (!d) return;
        var keys = this.visibleKeys(), i = _.indexOf(keys, this.selKey);
        i = Math.min(keys.length - 1, Math.max(0, i + d));
        keys.length && this.select(keys[i]);
        e.preventDefault();
        e.stopPropagation();
    },

    getFocusedComponents: function() {
        var arr = [this.inputSearch];
        this.chNoSpell && arr.push(this.chNoSpell);
        return arr.concat(this.getFooterButtons());
    },

    getDefaultFocusableComponent: function () {
        return this.inputSearch;
    },

    onBtnClick: function(event) {
        this.finish(event.currentTarget.attributes['result'].value);
    },

    onPrimary: function() {
        this.finish('ok');
        return false;
    },

    finish: function(result) {
        var rec = this.selected;
        if (result !== 'cancel' && rec) {
            this.rememberRecent(rec.value);
        }
        this.options.handler && this.options.handler.call(this, result, rec && rec.value, {code: rec && rec.code, noSpell: !!(this.chNoSpell && this.chNoSpell.isChecked())});
        this.close();
    },

    txtTitle        : 'Language',
    labelMark       : 'Mark selected text as:',
    labelSelect     : 'Select document language',
    txtSearch       : 'Search languages',
    txtCurrent      : 'Current language',
    txtRecent       : 'Recently used',
    txtAll          : 'All languages',
    txtNoResult     : 'No language found',
    txtSetDefault   : 'Set As Default',
    txtNoSpell      : 'Do not check spelling or grammar'
    }, Common.Views.LanguageDialog || {}))
});

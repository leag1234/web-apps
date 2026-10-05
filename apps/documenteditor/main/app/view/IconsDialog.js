/*
 * SPDX-FileCopyrightText: 2026 ElixirOffice contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 *
 * DE.Views.IconsDialog: Insert > Icons, like Word: categories, search, several
 * icons at once, a colour. The icon set comes with the theme
 * (common/main/resources/img/icons-gallery/icons.json, generated from Fluent UI
 * System Icons by the product repo); without it the Icons button is hidden.
 * The icons are inserted as pictures (PNG drawn from the vector paths).
 */

define([
    'common/main/lib/component/Window',
    'common/main/lib/component/InputField',
    'common/main/lib/component/ColorButton'
], function () {
    'use strict';

    var DATA_URL = '../../common/main/resources/img/icons-gallery/icons.json',
        dataPromise = null;

    DE.Views.IconsDialog = Common.UI.Window.extend(_.extend({
        options: {
            width: 640,
            cls: 'modal-dlg elx-icons-dlg'
        },

        // loads the icon set once; resolves to null when the theme has none
        loadData: function () {
            if (!dataPromise) {
                dataPromise = new Promise(function (resolve) {
                    $.ajax({url: DATA_URL, dataType: 'json', cache: true})
                        .done(function (data) { resolve(data && data.categories ? data : null); })
                        .fail(function () { resolve(null); });
                });
            }
            return dataPromise;
        },

        initialize: function (options) {
            _.extend(this.options, {title: this.txtTitle, buttons: [{value: 'ok', caption: this.txtInsert, primary: true}, 'cancel']}, options || {});
            this.template = [
                '<div class="box elx-icons-box">',
                    '<div class="elx-icons-top">',
                        '<div id="elx-icons-search" class="elx-icons-search"></div>',
                        '<label class="elx-icons-color-label">' + this.txtColor + '</label><div id="elx-icons-color"></div>',
                    '</div>',
                    '<div class="elx-icons-main">',
                        '<div class="elx-icons-cats" role="listbox"></div>',
                        '<div class="elx-icons-grid" role="listbox" aria-multiselectable="true"></div>',
                    '</div>',
                    '<div class="elx-icons-count"></div>',
                '</div>'
            ].join('');
            this.options.tpl = _.template(this.template)(this.options);
            this.handler = this.options.handler;
            this.data = this.options.data;
            this.selected = [];
            Common.UI.Window.prototype.initialize.call(this, this.options);
        },

        render: function () {
            Common.UI.Window.prototype.render.call(this);
            var me = this, $w = this.$window;
            this.inputSearch = new Common.UI.InputField({
                el: $w.find('#elx-icons-search'), allowBlank: true, validateOnChange: false,
                placeHolder: this.txtSearch, style: 'width: 100%;'
            });
            this.inputSearch.on('changing', function (input, value) { me.filter(value); });
            this.btnColor = new Common.UI.ColorButton({
                parentEl: $w.find('#elx-icons-color'), color: '000000', menu: true
            });
            this.btnColor.on('color:select', function () { me.renderGrid(); });
            var $cats = $w.find('.elx-icons-cats');
            _.each(this.data.categories, function (c, i) {
                $('<div class="elx-icons-cat" role="option" tabindex="0"></div>').text(me['txtCat_' + c.id] || c.id).attr('data-index', i).appendTo($cats);
            });
            $cats.on('click', '.elx-icons-cat', function () { me.showCategory(+$(this).attr('data-index')); });
            $w.find('.elx-icons-grid').on('click', '.elx-icons-item', function () { me.toggle($(this)); })
                .on('dblclick', '.elx-icons-item', function () { me.toggle($(this), true); me._handleInput('ok'); });
            $w.find('.dlg-btn').on('click', function () { me._handleInput($(this).attr('result')); });
            this.showCategory(0);
        },

        color: function () {
            var c = this.btnColor.color;
            if (typeof c === 'object' && c) c = c.color;
            return '#' + (c && c !== 'transparent' ? c : '000000');
        },

        svgOf: function (icon, size) {
            var fill = this.color();
            return '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" width="' + size + '" height="' + size + '">' +
                _.map(icon.d, function (d) { return '<path fill="' + fill + '" d="' + d + '"/>'; }).join('') + '</svg>';
        },

        showCategory: function (index) {
            this.inputSearch.setValue('');
            this.$window.find('.elx-icons-cat').removeClass('selected').eq(index).addClass('selected');
            this.list = this.data.categories[index].icons;
            this.renderGrid();
        },

        filter: function (text) {
            text = (text || '').trim().toLowerCase();
            if (!text) { this.showCategory(Math.max(0, this.$window.find('.elx-icons-cat.selected').index())); return; }
            var words = text.split(/\s+/), all = [];
            _.each(this.data.categories, function (c) { all = all.concat(c.icons); });
            this.$window.find('.elx-icons-cat').removeClass('selected');
            this.list = _.filter(all, function (ic) { return _.every(words, function (w) { return ic.n.indexOf(w) >= 0; }); });
            this.renderGrid();
        },

        renderGrid: function () {
            var me = this, $g = this.$window.find('.elx-icons-grid').empty();
            _.each(this.list, function (ic) {
                var sel = _.contains(me.selected, ic);
                $('<div class="elx-icons-item" role="option" tabindex="0"></div>').toggleClass('selected', sel).attr({'aria-selected': sel, title: ic.n.replace(/_/g, ' ')})
                    .html(me.svgOf(ic, 40)).data('icon', ic).appendTo($g);
            });
            this.updateCount();
        },

        toggle: function ($item, on) {
            var ic = $item.data('icon'), sel = on === true ? true : !_.contains(this.selected, ic);
            this.selected = sel ? _.union(this.selected, [ic]) : _.without(this.selected, ic);
            $item.toggleClass('selected', sel).attr('aria-selected', sel);
            this.updateCount();
        },

        updateCount: function () {
            var n = this.selected.length;
            this.$window.find('.elx-icons-count').text(n ? Common.Utils.String.format(this.txtSelected, n) : '');
            this.$window.find('.dlg-btn[result=ok]').prop('disabled', n === 0).toggleClass('disabled', n === 0);
        },

        // PNG pictures of the selected icons, drawn at print resolution
        getImages: function () {
            var me = this, px = 384;
            return Promise.all(_.map(this.selected, function (ic) {
                return new Promise(function (resolve) {
                    var img = new Image();
                    img.onload = function () {
                        var c = document.createElement('canvas');
                        c.width = c.height = px;
                        c.getContext('2d').drawImage(img, 0, 0, px, px);
                        resolve(c.toDataURL('image/png'));
                    };
                    img.onerror = function () { resolve(null); };
                    img.src = 'data:image/svg+xml;charset=utf-8,' + encodeURIComponent(me.svgOf(ic, px));
                });
            })).then(function (list) { return _.compact(list); });
        },

        _handleInput: function (state) {
            var me = this;
            if (state === 'ok') {
                if (!this.selected.length) return;
                this.getImages().then(function (urls) {
                    me.handler && me.handler.call(me, 'ok', urls);
                    me.close();
                });
                return;
            }
            this.handler && this.handler.call(this, state);
            this.close();
        },

        onPrimary: function () {
            this._handleInput('ok');
            return false;
        },

        txtTitle: 'Insert Icons',
        txtSearch: 'Search icons (English names)',
        txtColor: 'Color',
        txtInsert: 'Insert',
        txtSelected: '{0} selected',
        txtCat_arrows: 'Arrows',
        txtCat_people: 'People',
        txtCat_education: 'Education',
        txtCat_communication: 'Communication',
        txtCat_business: 'Business',
        txtCat_nature: 'Nature and weather',
        txtCat_animals: 'Animals',
        txtCat_food: 'Food and drink',
        txtCat_leisure: 'Leisure',
        txtCat_places: 'Places and transport',
        txtCat_symbols: 'Symbols and shapes',
        txtCat_technology: 'Technology'
    }, DE.Views.IconsDialog || {}));
});

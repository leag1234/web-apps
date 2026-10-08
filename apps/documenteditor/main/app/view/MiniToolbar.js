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
 *  MiniToolbar.js
 *
 *  Word's floating mini toolbar: font, size, grow/shrink, styles, Format
 *  Painter on the first row; bold, italic, underline, highlight, font colour,
 *  bullets, numbering and comment on the second. The controls reuse the
 *  toolbar's icons and strings; the controller (controller/MiniToolbar.js)
 *  positions it and wires it to the editor.
 */

define([
    'common/main/lib/component/Button',
    'common/main/lib/component/ColorButton',
    'common/main/lib/component/ComboBox',
    'common/main/lib/component/ComboBoxFonts',
    'common/main/lib/component/ThemeColorPalette'
], function () {
    'use strict';

    DE.Views.MiniToolbar = Backbone.View.extend(_.extend({
        el: null,

        initialize: function(options) {
            this.toolbar = options.toolbar;   // DE.Views.Toolbar, for its strings and its font store
        },

        render: function() {
            var me = this, tb = this.toolbar, T = DE.Views.Toolbar.prototype;
            var $el = $('<div id="mini-toolbar" class="mini-toolbar" role="toolbar" aria-label="' + me.textMiniToolbar + '">' +
                '<div class="mini-row"><span class="mini-slot" id="mini-elixir"></span><span class="mini-slot" id="mini-font-name"></span><span class="mini-slot" id="mini-font-size"></span>' +
                '<span class="mini-slot" id="mini-grow"></span><span class="mini-slot" id="mini-shrink"></span>' +
                '<span class="mini-slot" id="mini-styles"></span><span class="mini-slot" id="mini-painter"></span></div>' +
                '<div class="mini-row"><span class="mini-slot" id="mini-bold"></span><span class="mini-slot" id="mini-italic"></span>' +
                '<span class="mini-slot" id="mini-underline"></span><span class="mini-slot" id="mini-highlight"></span>' +
                '<span class="mini-slot" id="mini-fontcolor"></span><span class="mini-slot" id="mini-bullets"></span>' +
                '<span class="mini-slot" id="mini-numbering"></span><span class="mini-slot" id="mini-comment"></span></div>' +
                '</div>');
            $(document.body).append($el);
            this.$el = $el;
            this.el = $el[0];

            var button = function(slot, iconCls, hint, opts) {
                var btn = new Common.UI.Button(_.extend({
                    parentEl: $el.find('#' + slot),
                    cls: 'btn-toolbar',
                    iconCls: 'toolbar__icon ' + iconCls,
                    hint: hint
                }, opts || {}));
                return btn;
            };

            this.cmbFontName = new Common.UI.ComboBoxFonts({
                el: $el.find('#mini-font-name'),
                cls: 'input-group-nr',
                menuCls: 'scrollable-menu',
                menuStyle: 'min-width: 325px;',
                hint: T.tipFontName,
                store: tb && tb.cmbFontName ? tb.cmbFontName.store : new Common.Collections.Fonts(),
                recent: 0
            });
            this.cmbFontName.shareSprite(tb && tb.cmbFontName);
            this.cmbFontSize = new Common.UI.ComboBox({
                el: $el.find('#mini-font-size'),
                cls: 'input-group-nr',
                menuCls: 'scrollable-menu',
                menuStyle: 'min-width: 55px;max-height: 300px;',
                hint: T.tipFontSize,
                data: tb && tb.cmbFontSize ? tb.cmbFontSize.store.toJSON() : _.map([8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72], function(v) { return {value: v, displayValue: String(v)}; })
            });
            this.btnGrow = button('mini-grow', 'btn-incfont', T.tipIncFont);
            this.btnShrink = button('mini-shrink', 'btn-decfont', T.tipDecFont);
            this.btnStyles = button('mini-styles', 'btn-style', me.textStyles, {
                caption: me.textStyles,
                menu: new Common.UI.Menu({style: 'min-width: 180px;', maxHeight: 360, items: []})
            });
            this.btnPainter = button('mini-painter', 'btn-copystyle', T.tipCopyStyle, {enableToggle: true});

            this.btnBold = button('mini-bold', 'btn-bold', T.textBold, {enableToggle: true});
            this.btnItalic = button('mini-italic', 'btn-italic', T.textItalic, {enableToggle: true});
            this.btnUnderline = button('mini-underline', 'btn-underline', T.textUnderline, {enableToggle: true});
            this.btnHighlight = new Common.UI.ButtonColored({
                parentEl: $el.find('#mini-highlight'),
                cls: 'btn-toolbar',
                iconCls: 'toolbar__icon btn-highlight',
                hint: T.tipHighlightColor,
                split: true,
                menu: new Common.UI.Menu({
                    style: 'min-width: 100px;',
                    items: [
                        {template: _.template('<div id="mini-menu-highlight" style="width: 145px; display: inline-block;" class="palette-large"></div>')},
                        {caption: '--'},
                        this.mnuHighlightNone = new Common.UI.MenuItem({caption: T.strMenuNoFill})
                    ]
                })
            });
            this.btnFontColor = new Common.UI.ButtonColored({
                parentEl: $el.find('#mini-fontcolor'),
                cls: 'btn-toolbar',
                iconCls: 'toolbar__icon btn-fontcolor',
                hint: T.tipFontColor,
                split: true,
                menu: true,
                auto: true
            });
            this.btnBullets = button('mini-bullets', 'btn-setmarkers', T.tipMarkers, {enableToggle: true});
            this.btnNumbering = button('mini-numbering', 'btn-numbering', T.tipNumbers, {enableToggle: true});
            this.btnComment = button('mini-comment', 'btn-add-comment', me.textNewComment);
            // Elixir AI: the "Ask Elixir" menu (suggestions for the selection and a free-text field)
            var AI = window.Common && Common.ElixirAI;
            this.btnElixir = new Common.UI.Button({
                parentEl: $el.find('#mini-elixir'),
                cls: 'btn-toolbar mini-elixir',
                iconCls: 'elixir-ai-icon elixir-ai-icon-sparkle',
                hint: AI && AI.t ? AI.t('menuAsk') : 'Ask Elixir'
            });

            // colour pickers
            this.btnHighlight.menu.on('show:after', function() {
                if (!me.mnuHighlightPicker) {
                    me.mnuHighlightPicker = new Common.UI.ThemeColorPalette({
                        el: $('#mini-menu-highlight'),
                        colors: ['FFFF00', '00FF00', '00FFFF', 'FF00FF', '0000FF', 'FF0000', '00008B', '008B8B',
                            '006400', '800080', '8B0000', '808000', 'FFFFFF', 'D3D3D3', 'A9A9A9', '000000'],
                        value: 'FFFF00', dynamiccolors: 0, themecolors: 0, effects: 0, columns: 4
                    });
                    me.mnuHighlightPicker.on('select', function(picker, color) { me.trigger('highlight:select', color); });
                }
            });
            this.btnFontColor.setMenu();
            this.mnuFontColorPicker = this.btnFontColor.getPicker();
            this.btnFontColor.menu.on('show:after', function() {
                me.mnuFontColorPicker.updateColors(Common.Utils.ThemeColor.getEffectColors(), Common.Utils.ThemeColor.getStandartColors());
            });
            return this;
        },

        textMiniToolbar: 'Mini toolbar',
        textStyles: 'Styles',
        textNewComment: 'New Comment'

    }, DE.Views.MiniToolbar || {}));
});

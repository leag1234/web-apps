/*
 * SPDX-FileCopyrightText: 2026 ElixirOffice contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 *
 * Word's Paragraph dialog on top of the paragraph settings of the editor (Paragraph group
 * launcher). DE.Views.ParagraphSettingsAdvanced keeps all the logic; only the layout changes:
 *  - DE.Views.WordParagraphDialog: two tabs as in Word, "Indents and Spacing" (General,
 *    Indentation, Spacing, Preview) and "Line and Page Breaks" (Pagination, Formatting
 *    exceptions, Preview), with a "Tabs..." button that applies the dialog and opens the tabs;
 *  - DE.Views.WordTabsDialog: the tab stops of the same dialog alone, as Word's Tabs dialog.
 * The other panels of the base template stay in the DOM (hidden) so the base code finds its
 * controls.
 */

define([
    'text!documenteditor/main/app/template/ParagraphSettingsAdvanced.template',
    'documenteditor/main/app/view/ParagraphSettingsAdvanced'
], function (baseTemplate) { 'use strict';

    var Base = DE.Views.ParagraphSettingsAdvanced;
    var REST_ANCHOR = '<div id="id-adv-paragraph-borders"';

    var wordPanels = function (s) {
        var preview = '<div class="word-dlg-section">' + s.txtPreview + '</div>' +
            '<div class="word-para-preview"><div class="pv-other"></div><div class="pv-cur"></div><div class="pv-other"></div></div>';
        return [
            '<div id="id-adv-paragraph-indents" class="settings-panel active">',
                '<div class="inner-content word-para">',
                    '<div class="word-dlg-section">' + s.txtGeneral + '</div>',
                    '<table class="word-para-grid" role="presentation"><tbody>',
                        '<tr><td class="lbl"><label>' + s.txtAlignment + '</label></td><td><div id="paragraphadv-spin-text-alignment"></div></td>',
                            '<td class="lbl text-only"><label>' + s.txtOutlineLevel + '</label></td><td class="text-only"><div id="paragraphadv-spin-outline-level"></div></td></tr>',
                        '<tr><td class="lbl"><label>' + s.txtDirection + '</label></td><td colspan="3"><div id="paragraphadv-dir-ltr" class="word-para-radio"></div><div id="paragraphadv-dir-rtl" class="word-para-radio"></div></td></tr>',
                    '</tbody></table>',
                    '<div class="word-dlg-section">' + s.txtIndentation + '</div>',
                    '<table class="word-para-grid" role="presentation"><tbody>',
                        '<tr><td class="lbl"><label id="paragraphadv-lbl-indent-left">' + s.txtLeft + '</label></td><td><div id="paragraphadv-spin-indent-left"></div></td>',
                            '<td class="lbl-top"><label>' + s.txtSpecial + '</label></td><td class="lbl-top"><label>' + s.txtBy + '</label></td></tr>',
                        '<tr><td class="lbl"><label id="paragraphadv-lbl-indent-right">' + s.txtRight + '</label></td><td><div id="paragraphadv-spin-indent-right"></div></td>',
                            '<td><div id="paragraphadv-spin-special"></div></td><td><div id="paragraphadv-spin-special-by"></div></td></tr>',
                    '</tbody></table>',
                    '<div class="word-dlg-section">' + s.txtSpacing + '</div>',
                    '<table class="word-para-grid" role="presentation"><tbody>',
                        '<tr><td class="lbl"><label>' + s.txtBefore + '</label></td><td><div id="paragraphadv-spin-spacing-before"></div></td>',
                            '<td class="lbl-top"><label>' + s.txtLineSpacing + '</label></td><td class="lbl-top"><label>' + s.txtAt + '</label></td></tr>',
                        '<tr><td class="lbl"><label>' + s.txtAfter + '</label></td><td><div id="paragraphadv-spin-spacing-after"></div></td>',
                            '<td><div id="paragraphadv-spin-line-rule"></div></td><td><div id="paragraphadv-spin-line-height"></div></td></tr>',
                        '<tr class="text-only"><td colspan="4"><div id="paragraphadv-checkbox-add-interval"></div></td></tr>',
                    '</tbody></table>',
                    '<div style="display: none;"><div id="paragraphadv-indent-preview"></div></div>',
                    preview,
                '</div>',
            '</div>',
            '<div id="id-adv-paragraph-line" class="settings-panel">',
                '<div class="inner-content word-para text-only">',
                    '<div class="word-dlg-section">' + s.txtPagination + '</div>',
                    '<div class="word-para-checks">',
                        '<div id="paragraphadv-checkbox-orphan"></div>',
                        '<div id="paragraphadv-checkbox-keep-next"></div>',
                        '<div id="paragraphadv-checkbox-keep-lines"></div>',
                        '<div id="paragraphadv-checkbox-break-before"></div>',
                    '</div>',
                    '<div class="word-dlg-section">' + s.txtExceptions + '</div>',
                    '<div class="word-para-checks"><div id="paragraphadv-checkbox-suppress-line-numbers"></div></div>',
                    preview,
                '</div>',
            '</div>'
        ].join('');
    };

    var PX_PER_MM = 2;

    DE.Views.WordParagraphDialog = Base.extend(_.extend({
        options: _.extend({}, Base.prototype.options, {
            contentWidth: 470,
            contentHeight: 470,
            toggleGroup: 'word-paragraph-dlg-group',
            storageName: 'de-word-paragraph-dlg-tab'
        }),

        initialize: function (options) {
            var rest = baseTemplate.substring(baseTemplate.indexOf(REST_ANCHOR));
            options = _.extend({}, options);
            options.contentTemplate = _.template(wordPanels(this) + rest)({scope: this});
            options.buttons = [{value: 'tabs', caption: this.txtTabsBtn}, 'ok', 'cancel'];
            options.title = this.txtTitle;
            Base.prototype.initialize.call(this, options);
        },

        render: function () {
            Base.prototype.render.call(this);
            var me = this, $w = this.getChild();
            $w.addClass('word-paragraph-dlg word-dlg-top-tabs');
            this.btnsCategory[0].setCaption(this.txtIndentsTab);
            this.btnsCategory[1].setCaption(this.txtBreaksTab);
            for (var i = 2; i < this.btnsCategory.length; i++)
                this.btnsCategory[i].setVisible(false);
            $w.find('.footer button[result="tabs"]').addClass('word-para-tabs-btn');
            // the category list is now a tab strip above the content: the window is as wide as the content
            $w.width(this.content_panel.outerWidth() + 2);
            this.on('show', function () { $w.width(me.content_panel.outerWidth() + 2); });

            this.chOrphan.setCaption && this.chOrphan.setCaption(this.txtWidow);
            var upd = function () { me.updateWordPreview(); };
            _.each([this.numIndentsLeft, this.numIndentsRight, this.numSpecialBy, this.numSpacingBefore, this.numSpacingAfter, this.numLineHeight], function (c) { c && c.on('change', upd); });
            _.each([this.cmbTextAlignment, this.cmbSpecial, this.cmbLineRule], function (c) { c && c.on('selected', upd); });
            this.rbDirRtl && this.rbDirRtl.on('change', upd);
            this.on('show', upd);
            _.defer(upd);
        },

        setActiveCategory: function (index) {
            Base.prototype.setActiveCategory.call(this, Math.min(parseInt(index) || 0, 1));
        },

        updateWordPreview: function () {
            var $w = this.getChild(), M = Common.Utils.Metric;
            if (!this.numIndentsLeft) return;
            var mm = function (spin) { var v = spin ? spin.getNumberValue() : 0; return (typeof v === 'number' && !isNaN(v)) ? M.fnRecalcToMM(v) : 0; };
            var left = mm(this.numIndentsLeft), right = mm(this.numIndentsRight), by = mm(this.numSpecialBy),
                before = mm(this.numSpacingBefore), after = mm(this.numSpacingAfter),
                special = this.cmbSpecial ? this.cmbSpecial.getValue() : 0,
                align = this.cmbTextAlignment ? this.cmbTextAlignment.getValue() : 1,
                rule = this.cmbLineRule ? this.cmbLineRule.getValue() : 1,
                line = this.numLineHeight ? this.numLineHeight.getNumberValue() : 1,
                rtl = this.rbDirRtl && this.rbDirRtl.getValue();
            var css = {
                'margin-left': Math.min(Math.max(left, 0) * PX_PER_MM, 200) + 'px',
                'margin-right': Math.min(Math.max(right, 0) * PX_PER_MM, 200) + 'px',
                'text-indent': (special === 1 ? by : (special === 2 ? -by : 0)) * PX_PER_MM + 'px',
                'padding-left': special === 2 ? by * PX_PER_MM + 'px' : '0',
                'margin-top': Math.min(before * PX_PER_MM, 30) + 'px',
                'margin-bottom': Math.min(after * PX_PER_MM, 30) + 'px',
                'text-align': ({0: 'right', 1: 'left', 2: 'center', 3: 'justify'})[align] || 'left',
                'line-height': rule === 1 ? Math.max(1, Math.min(line || 1, 3)) * 9 + 'px' : '9px',
                'direction': rtl ? 'rtl' : 'ltr'
            };
            var sample = _.times(14, _.constant(this.txtSampleText)).join(' '),
                other = _.times(18, _.constant(this.txtOtherText)).join(' ');
            $w.find('.word-para-preview .pv-cur').text(sample).css(css);
            $w.find('.word-para-preview .pv-other').text(other);
        },

        txtTitle: 'Paragraph',
        txtIndentsTab: 'Indents and Spacing',
        txtBreaksTab: 'Line and Page Breaks',
        txtTabsBtn: 'Tabs...',
        txtGeneral: 'General',
        txtAlignment: 'Alignment:',
        txtOutlineLevel: 'Outline level:',
        txtDirection: 'Direction:',
        txtIndentation: 'Indentation',
        txtLeft: 'Left:',
        txtRight: 'Right:',
        txtSpecial: 'Special:',
        txtBy: 'By:',
        txtSpacing: 'Spacing',
        txtBefore: 'Before:',
        txtAfter: 'After:',
        txtLineSpacing: 'Line spacing:',
        txtAt: 'At:',
        txtPreview: 'Preview',
        txtPagination: 'Pagination',
        txtWidow: 'Widow/Orphan control',
        txtExceptions: 'Formatting exceptions',
        txtSampleText: 'Sample Text',
        txtOtherText: 'Previous Paragraph'
    }, DE.Views.WordParagraphDialog || {}));

    DE.Views.WordTabsDialog = Base.extend(_.extend({
        options: _.extend({}, Base.prototype.options, {
            contentWidth: 370,
            contentHeight: 300,
            toggleGroup: 'word-tabs-dlg-group',
            storageName: null
        }),

        initialize: function (options) {
            options = _.extend({}, options, {title: this.txtTitle});
            Base.prototype.initialize.call(this, options);
        },

        render: function () {
            Base.prototype.render.call(this);
            var $w = this.getChild();
            $w.addClass('word-tabs-dlg');
            this.setActiveCategory(4);
            $w.find('.menu-panel, .box > .separator:not(.horizontal)').hide();
            $w.width(this.content_panel.outerWidth() + 2);
        },

        txtTitle: 'Tabs'
    }, DE.Views.WordTabsDialog || {}));
});

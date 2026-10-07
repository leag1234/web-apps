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

define([
    'common/main/lib/component/Window',
    'common/main/lib/component/RadioBox'
], function () { 'use strict';

    // Word's Paste Special: how the clipboard is pasted. The browser clipboard gives one rich entry
    // (HTML) and, for images, a picture: Formatted text keeps the source formatting.
    DE.Views.PasteSpecialDialog = Common.UI.Window.extend(_.extend({
        options: {
            width: 380,
            header: true,
            cls: 'modal-dlg',
            id: 'window-paste-special',
            buttons: ['ok', 'cancel']
        },

        initialize : function(options) {
            _.extend(this.options, {title: this.textTitle}, options || {});
            this.template = [
                '<div class="box" style="padding: 0 4px;">',
                    '<label class="input-label" style="display:block;margin-bottom:8px;">' + this.textPasteAs + '</label>',
                    '<div id="pspecial-formatted" style="margin-bottom: 8px;"></div>',
                    '<div id="pspecial-merge" style="margin-bottom: 8px;"></div>',
                    '<div id="pspecial-text" style="margin-bottom: 8px;"></div>',
                    '<div id="pspecial-picture" style="margin-bottom: 8px;"></div>',
                '</div>'
            ].join('');
            this.options.tpl = _.template(this.template)(this.options);
            Common.UI.Window.prototype.initialize.call(this, this.options);
        },

        render: function() {
            Common.UI.Window.prototype.render.call(this);
            var P = Asc.c_oSpecialPasteProps, $w = this.getChild(), me = this;
            this.radios = [];
            _.each([
                ['#pspecial-formatted', this.textFormatted, P.sourceformatting, true],
                ['#pspecial-merge', this.textMerge, P.mergeFormatting],
                ['#pspecial-text', this.textUnformatted, P.keepTextOnly],
                ['#pspecial-picture', this.textPicture, P.picture]
            ], function(r) {
                me.radios.push({value: r[2], rb: new Common.UI.RadioBox({el: $w.find(r[0]), labelText: r[1], name: 'asc-radio-paste-special', checked: !!r[3]})});
            });
            $w.find('.dlg-btn').on('click', _.bind(this.onBtnClick, this));
            return this;
        },

        getFocusedComponents: function() {
            return _.pluck(this.radios, 'rb').concat(this.getFooterButtons());
        },

        getDefaultFocusableComponent: function() {
            return this.radios[0].rb;
        },

        onPrimary: function() {
            this.finish('ok');
            return false;
        },

        onBtnClick: function(event) {
            this.finish(event.currentTarget.attributes['result'].value);
        },

        finish: function(result) {
            var chosen = _.find(this.radios, function(r) { return r.rb.getValue(); });
            this.options.handler && this.options.handler.call(this, result, chosen ? chosen.value : null);
            this.close();
        },

        textTitle: 'Paste Special',
        textPasteAs: 'Paste as:',
        textFormatted: 'Formatted text (keep source formatting)',
        textMerge: 'Merge formatting',
        textUnformatted: 'Unformatted text',
        textPicture: 'Picture'
    }, DE.Views.PasteSpecialDialog || {}));
});

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
 *  GoToDialog.js
 *
 *  Word's "Go To" (Ctrl+G, F5): page, bookmark, footnote or endnote.
 */

define([
    'common/main/lib/component/Window',
    'common/main/lib/component/ComboBox',
    'common/main/lib/component/InputField'
], function () { 'use strict';

    DE.Views.GoToDialog = Common.UI.Window.extend(_.extend({
        options: {
            width: 420,
            header: true,
            style: 'min-width: 360px;',
            cls: 'modal-dlg',
            id: 'window-goto',
            buttons: null
        },

        initialize : function(options) {
            _.extend(this.options, {title: this.textTitle}, options || {});
            this.api = this.options.api;

            this.template = [
                '<div class="box" style="padding: 0 4px;">',
                    '<div style="display: flex; gap: 16px;">',
                        '<div style="width: 140px;">',
                            '<label class="input-label">' + this.textGoToWhat + '</label>',
                            '<div id="goto-what" style="width: 100%;"></div>',
                        '</div>',
                        '<div style="flex: 1;">',
                            '<label id="goto-value-label" class="input-label">' + this.textEnterPage + '</label>',
                            '<div id="goto-value" style="width: 100%;"></div>',
                            '<div id="goto-bookmark" style="width: 100%; display: none;"></div>',
                        '</div>',
                    '</div>',
                    '<div id="goto-hint" class="light-text" style="margin-top: 10px; font-size: 11px;">' + this.textRelativeHint + '</div>',
                '</div>',
                '<div class="footer right">',
                    '<button class="btn normal dlg-btn" result="prev" style="min-width: 86px;">' + this.textPrevious + '</button>',
                    '<button class="btn normal dlg-btn primary" result="goto" style="min-width: 86px;">' + this.textGoTo + '</button>',
                    '<button class="btn normal dlg-btn" result="cancel" style="min-width: 86px;">' + this.textClose + '</button>',
                '</div>'
            ].join('');
            this.options.tpl = _.template(this.template)(this.options);

            Common.UI.Window.prototype.initialize.call(this, this.options);
        },

        render: function() {
            Common.UI.Window.prototype.render.call(this);
            var me = this, $window = this.getChild();

            this.cmbWhat = new Common.UI.ComboBox({
                el: $window.find('#goto-what'),
                cls: 'input-group-nr',
                style: 'width: 100%;',
                menuStyle: 'min-width: 100%;',
                editable: false,
                data: [
                    {value: 'page', displayValue: this.textPage},
                    {value: 'bookmark', displayValue: this.textBookmark},
                    {value: 'footnote', displayValue: this.textFootnote},
                    {value: 'endnote', displayValue: this.textEndnote}
                ],
                takeFocusOnClose: true
            });
            this.cmbWhat.setValue('page');
            this.cmbWhat.on('selected', _.bind(this.onWhatChanged, this));

            this.inputValue = new Common.UI.InputField({
                el: $window.find('#goto-value'),
                style: 'width: 100%;',
                validateOnBlur: false
            });
            this.inputValue.on('changing', _.bind(this.updateButtons, this));

            var bookmarks = [];
            var manager = this.api && this.api.asc_GetBookmarksManager ? this.api.asc_GetBookmarksManager() : null;
            if (manager) {
                for (var i = 0, n = manager.asc_GetCount(); i < n; i++) {
                    var name = manager.asc_GetName(i);
                    !manager.asc_IsInternalUseBookmark(name) && bookmarks.push({value: name, displayValue: name});
                }
            }
            this.cmbBookmark = new Common.UI.ComboBox({
                el: $window.find('#goto-bookmark'),
                cls: 'input-group-nr',
                style: 'width: 100%;',
                menuStyle: 'min-width: 100%; max-height: 200px;',
                editable: true,
                data: bookmarks,
                takeFocusOnClose: true
            });
            bookmarks.length && this.cmbBookmark.setValue(bookmarks[0].value);

            $window.find('.dlg-btn').on('click', _.bind(this.onBtnClick, this));
            this.updateButtons();
            return this;
        },

        show: function() {
            Common.UI.Window.prototype.show.apply(this, arguments);
            var me = this;
            _.delay(function() { me.inputValue.focus(); }, 50);
        },

        getFocusedComponents: function() {
            return [this.cmbWhat, this.inputValue, this.cmbBookmark].concat(this.getFooterButtons ? this.getFooterButtons() : []);
        },

        getDefaultFocusableComponent: function() {
            return this.inputValue;
        },

        onWhatChanged: function() {
            var what = this.cmbWhat.getValue(), $window = this.getChild();
            var isBookmark = what === 'bookmark';
            $window.find('#goto-value').toggle(!isBookmark);
            $window.find('#goto-bookmark').toggle(isBookmark);
            $window.find('#goto-hint').toggle(what !== 'bookmark');
            $window.find('#goto-value-label').text(isBookmark ? this.textEnterBookmark : (what === 'page' ? this.textEnterPage : (what === 'footnote' ? this.textEnterFootnote : this.textEnterEndnote)));
            this.updateButtons();
        },

        updateButtons: function() {
            var what = this.cmbWhat.getValue(), $window = this.getChild();
            var empty = what !== 'bookmark' && !(this.inputValue.getValue() || '').trim();
            // like Word: "Next" while the box is empty, "Go To" once something is typed
            $window.find('.dlg-btn[result="goto"]').text(empty && what !== 'bookmark' ? this.textNext : this.textGoTo);
            $window.find('.dlg-btn[result="prev"]').prop('disabled', what === 'bookmark');
        },

        onBtnClick: function(event) {
            var result = $(event.currentTarget).attr('result');
            if (result === 'cancel') {
                this.close();
                return;
            }
            this.go(result === 'prev' ? -1 : 0);
        },

        onPrimary: function() {
            this.go(0);
            return false;
        },

        // dir: -1 previous, 0 "Go To" (or "Next" when the box is empty)
        go: function(dir) {
            var api = this.api;
            if (!api) return;
            var what = this.cmbWhat.getValue();
            var text = (this.inputValue.getValue() || '').trim();

            if (what === 'bookmark') {
                var name = this.cmbBookmark.getValue();
                var manager = api.asc_GetBookmarksManager();
                name && manager && manager.asc_HaveBookmark(name) && manager.asc_GoToBookmark(name);
                return;
            }

            if (what === 'page') {
                var count = api.getCountPages(), cur = api.getCurrentPage();
                var target;
                if (dir < 0) target = cur - 1;
                else if (!text) target = cur + 1;
                else if (/^[+-]\d+$/.test(text)) target = cur + parseInt(text, 10);
                else if (/^\d+$/.test(text)) target = parseInt(text, 10) - 1;
                else return;
                target = Math.max(0, Math.min(count - 1, target));
                api.goToPage(target);
                return;
            }

            var gotoNote = what === 'footnote' ? 'asc_GotoFootnote' : 'asc_GotoEndnote';
            if (dir < 0)
                api[gotoNote](false);
            else if (!text || /^\+\d+$/.test(text))
                for (var i = 0, n = text ? parseInt(text, 10) : 1; i < n; i++) api[gotoNote](true);
            else if (/^-\d+$/.test(text))
                for (var j = 0, m = -parseInt(text, 10); j < m; j++) api[gotoNote](false);
            else if (/^\d+$/.test(text)) {
                api.asc_MoveCursorToStartPos ? api.asc_MoveCursorToStartPos() : null;
                for (var k = 0, c = parseInt(text, 10); k < c; k++) api[gotoNote](true);
            }
        },

        textTitle: 'Go To',
        textGoToWhat: 'Go to what:',
        textPage: 'Page',
        textBookmark: 'Bookmark',
        textFootnote: 'Footnote',
        textEndnote: 'Endnote',
        textEnterPage: 'Enter page number:',
        textEnterBookmark: 'Enter bookmark name:',
        textEnterFootnote: 'Enter footnote number:',
        textEnterEndnote: 'Enter endnote number:',
        textRelativeHint: 'Enter + and - to move relative to the current location. Example: +4 will move forward four items.',
        textPrevious: 'Previous',
        textNext: 'Next',
        textGoTo: 'Go To',
        textClose: 'Close'

    }, DE.Views.GoToDialog || {}));
});

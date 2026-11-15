// SPDX-License-Identifier: AGPL-3.0-only
/**
 * DE.Controllers.Sort: Home > Paragraph > Sort and Table Layout > Data > Sort.
 * Opens the Sort dialog for the selected paragraphs or table rows; the sorting itself is done by the
 * editor (asc_GetSortInfo / asc_Sort), in one undo step.
 */
define([
    'core',
    'documenteditor/main/app/view/SortDialog'
], function () {
    'use strict';

    DE.Controllers.Sort = Backbone.Controller.extend(_.extend({
        initialize: function () {},

        onLaunch: function () {
            Common.NotificationCenter.on('app:ready', _.bind(function () {
                this.api = this.getApplication().getController('Viewport').getApi();
            }, this));
        },

        open: function () {
            var api = this.api || this.getApplication().getController('Viewport').getApi(), me = this;
            if (!api || !api.asc_GetSortInfo) return;
            var info = api.asc_GetSortInfo({separator: 'tab', separatorChar: ''});
            if (!info) return;
            if (info.error) {
                this.showError(info.error);
                return;
            }
            var lang = (Common.Locale && Common.Locale.getCurrentLanguage && Common.Locale.getCurrentLanguage()) || 'en';
            (new DE.Views.SortDialog({
                api: api,
                info: info,
                lang: lang,
                handler: function (dlg, result) {
                    if (result == 'ok') {
                        var code = api.asc_Sort(dlg.getSettings());
                        if (code !== 'ok') {
                            me.showError(code);
                            return;
                        }
                    }
                    Common.NotificationCenter.trigger('edit:complete');
                }
            })).show();
        },

        showError: function (code) {
            var msg = DE.Controllers.Sort.errorText(this, code);
            Common.UI.warning({
                title: this.txtTitle,
                msg: msg,
                maxwidth: 520,
                callback: function () { Common.NotificationCenter.trigger('edit:complete'); }
            });
        },

        txtTitle: 'Sort',
        txtErrNone: 'Select at least two paragraphs, or put the cursor in a table with at least two rows, to sort.',
        txtErrMerged: 'This table cannot be sorted because it has merged cells. Unmerge the cells and try again.',
        txtErrNested: 'This table cannot be sorted because it contains a nested table.',
        txtErrSection: 'The selection contains a section break. Sort the paragraphs on each side of it separately.',
        txtErrMixed: 'The selection cannot be sorted. Select consecutive paragraphs in one place, or rows of a single table.',
        txtErrLocked: 'The selection is protected or locked, so it cannot be sorted.'
    }, DE.Controllers.Sort || {}));

    // message for an error code returned by the editor
    DE.Controllers.Sort.errorText = function (strings, code) {
        var map = {none: 'txtErrNone', merged: 'txtErrMerged', nested: 'txtErrNested', section: 'txtErrSection', mixed: 'txtErrMixed', locked: 'txtErrLocked'};
        return strings[map[code] || 'txtErrMixed'];
    };

    return DE.Controllers.Sort;
});

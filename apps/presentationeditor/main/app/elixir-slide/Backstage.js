/*
 * SPDX-FileCopyrightText: 2026 ElixirOffice contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 *
 * ElixirOffice window chrome: PowerPoint-like backstage (presentation editor).
 * The backstage is shared (common/main/lib/elixir-chrome/Backstage.js); this
 * subclass gives PowerPoint's file formats and presentation wording, and no
 * Restrict Editing (presentations have no document protection).
 * Translations: locale PE.Views.ElixirBackstage.*.
 */
define([
    'core',
    'common/main/lib/elixir-chrome/Backstage'
], function () {
    'use strict';

    var STRINGS = {
        txtBlankDocument: 'Blank presentation', txtDocumentTypes: 'Presentation file types',
        txtPdfBullet1: 'Preserves the layout, fonts and images of every slide',
        txtFmtPptx: 'PowerPoint presentation (*.pptx)', txtFmtPptxDesc: 'The standard format, opens everywhere',
        txtFmtPptm: 'Macro-enabled presentation (*.pptm)', txtFmtPptmDesc: 'Keeps the macros of the presentation',
        txtFmtPotx: 'Presentation template (*.potx)', txtFmtPotxDesc: 'Starting point for new presentations',
        txtFmtPpsx: 'PowerPoint show (*.ppsx)', txtFmtPpsxDesc: 'Opens directly as a slide show',
        txtFmtOdp: 'OpenDocument presentation (*.odp)', txtFmtOdpDesc: 'Open standard, for LibreOffice and others',
        txtFmtOtp: 'OpenDocument presentation template (*.otp)', txtFmtOtpDesc: 'Open standard template',
        txtFmtImgDesc: 'One image per slide, in a zip file'
    };

    PE.Views.ElixirBackstage = Common.Views.ElixirBackstage.extend(_.extend({
        defaultFileType: 'pptx',

        formats: function () {
            var T = Asc.c_oAscFileType;
            return [
                { type: T.PPTX, ext: '.pptx', icon: 'slides', name: 'txtFmtPptx', desc: 'txtFmtPptxDesc', group: 'doc' },
                { type: T.PPTM, ext: '.pptm', icon: 'code', name: 'txtFmtPptm', desc: 'txtFmtPptmDesc', group: 'doc', onlyFor: 'pptm' },
                { type: T.POTX, ext: '.potx', icon: 'template', name: 'txtFmtPotx', desc: 'txtFmtPotxDesc', group: 'doc' },
                { type: T.PPSX, ext: '.ppsx', icon: 'slideshow', name: 'txtFmtPpsx', desc: 'txtFmtPpsxDesc', group: 'doc' },
                { type: T.ODP, ext: '.odp', icon: 'slides', name: 'txtFmtOdp', desc: 'txtFmtOdpDesc', group: 'doc' },
                { type: T.OTP, ext: '.otp', icon: 'template', name: 'txtFmtOtp', desc: 'txtFmtOtpDesc', group: 'doc' },
                { type: T.PDF, ext: '.pdf', icon: 'pdf', name: 'txtFmtPdf', desc: 'txtFmtPdfDesc', group: 'pdf' },
                { type: T.PDFA, ext: '.pdf', icon: 'pdf', name: 'txtFmtPdfa', desc: 'txtFmtPdfaDesc', group: 'pdf' },
                { type: T.PNG, ext: '.zip', icon: 'image', name: 'txtFmtPng', desc: 'txtFmtImgDesc', group: 'image' },
                { type: T.JPG, ext: '.zip', icon: 'image', name: 'txtFmtJpg', desc: 'txtFmtImgDesc', group: 'image' }
            ];
        },

        defaultFormat: function () { return Asc.c_oAscFileType.PPTX; },

        // no document protection in presentations (password and signature stay in Protect)
        canRestrictEditing: function () { return false; }
    }, STRINGS, PE.Views.ElixirBackstage || {}));
});

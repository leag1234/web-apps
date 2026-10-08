/*
 * SPDX-FileCopyrightText: 2026 ElixirOffice contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 *
 * Word's Design tab presets (document editor): themes (color scheme + theme fonts), theme
 * font pairs, style sets and paragraph spacing. The values follow Word's own presets; the
 * themes are the Office themes whose color schemes the engine ships.
 */
define([], function () {
    'use strict';

    var P = window.ElixirDesignPresets = {};

    // Design > Themes: [theme, color scheme of the engine, heading font, body font]. The Office themes
    // and the themes whose color scheme the engine ships under the same name (Aspect, Median, Paper,
    // Marquee, Slipstream) keep Word's fonts; the others are named after their color scheme.
    P.themes = [
        ['Office', 'Office', 'Aptos Display', 'Aptos'],
        ['Office 2013 - 2022', 'Office 2013 - 2022', 'Calibri Light', 'Calibri'],
        ['Office 2007 - 2010', 'Office 2007 - 2010', 'Cambria', 'Calibri'],
        ['Aspect', 'Aspect', 'Verdana', 'Verdana'],
        ['Median', 'Median', 'Tw Cen MT', 'Tw Cen MT'],
        ['Paper', 'Paper', 'Constantia', 'Constantia'],
        ['Marquee', 'Marquee', 'Tw Cen MT Condensed', 'Tw Cen MT'],
        ['Slipstream', 'Slipstream', 'Trebuchet MS', 'Trebuchet MS'],
        ['Blue', 'Blue', 'Calibri Light', 'Calibri'],
        ['Blue II', 'Blue II', 'Century Gothic', 'Century Gothic'],
        ['Blue Warm', 'Blue Warm', 'Corbel', 'Corbel'],
        ['Blue Green', 'Blue Green', 'Franklin Gothic Medium', 'Franklin Gothic Book'],
        ['Green', 'Green', 'Trebuchet MS', 'Trebuchet MS'],
        ['Green Yellow', 'Green Yellow', 'Garamond', 'Garamond'],
        ['Yellow', 'Yellow', 'Gill Sans MT', 'Gill Sans MT'],
        ['Yellow Orange', 'Yellow Orange', 'Tw Cen MT', 'Tw Cen MT'],
        ['Orange', 'Orange', 'Calibri Light', 'Calibri'],
        ['Orange Red', 'Orange Red', 'Century Gothic', 'Century Gothic'],
        ['Red', 'Red', 'Gill Sans MT', 'Gill Sans MT'],
        ['Red Violet', 'Red Violet', 'Georgia', 'Georgia'],
        ['Violet', 'Violet', 'Corbel', 'Corbel'],
        ['Grayscale', 'Grayscale', 'Arial', 'Arial']
    ];

    // Design > Fonts: [name, heading font, body font]
    P.fonts = [
        ['Office', 'Aptos Display', 'Aptos'],
        ['Office 2013 - 2022', 'Calibri Light', 'Calibri'],
        ['Office 2007 - 2010', 'Cambria', 'Calibri'],
        ['Calibri', 'Calibri', 'Calibri'],
        ['Arial', 'Arial', 'Arial'],
        ['Corbel', 'Corbel', 'Corbel'],
        ['Candara', 'Candara', 'Candara'],
        ['Franklin Gothic', 'Franklin Gothic Medium', 'Franklin Gothic Book'],
        ['Century Gothic', 'Century Gothic', 'Century Gothic'],
        ['Tw Cen MT', 'Tw Cen MT', 'Tw Cen MT'],
        ['Cambria', 'Cambria', 'Cambria'],
        ['Garamond', 'Garamond', 'Garamond'],
        ['Georgia', 'Georgia', 'Georgia'],
        ['Times New Roman', 'Times New Roman', 'Times New Roman'],
        ['Trebuchet MS', 'Trebuchet MS', 'Trebuchet MS'],
        ['Verdana', 'Verdana', 'Verdana'],
        ['Liberation', 'Liberation Sans', 'Liberation Serif']
    ];

    // Design > Paragraph Spacing: [key, before pt, after pt, line spacing (multiple)]
    P.spacing = [
        ['default', 0, 8, 1.08],
        ['none', 0, 0, 1],
        ['compact', 0, 4, 1],
        ['tight', 0, 6, 1.15],
        ['open', 0, 10, 1.15],
        ['relaxed', 0, 6, 1.5],
        ['double', 0, 8, 2]
    ];

    // Design > Style Sets. Per style: text (size pt, bold, italic, caps, color hex or null = theme text,
    // accent = theme accent 1), paragraph (align, before/after pt, bottom border hex, shading hex).
    // The fonts stay the theme fonts.
    P.styleSets = [
        { key: 'basic', name: 'Basic (Simple)',
          Title: { size: 28, caps: false, align: 'left', after: 0, color: null },
          'Heading 1': { size: 16, bold: false, color: 'accent', before: 12, after: 0 },
          'Heading 2': { size: 13, bold: false, color: 'accent', before: 2, after: 0 } },
        { key: 'elegant', name: 'Basic (Elegant)',
          Title: { size: 28, caps: true, align: 'center', after: 12, color: null, border: '404040' },
          'Heading 1': { size: 14, caps: true, bold: false, color: null, before: 18, after: 6 },
          'Heading 2': { size: 12, caps: true, bold: false, color: '595959', before: 10, after: 4 } },
        { key: 'stylish', name: 'Basic (Stylish)',
          Title: { size: 26, bold: true, align: 'left', after: 6, color: 'accent' },
          'Heading 1': { size: 14, bold: true, color: 'accent', before: 18, after: 4 },
          'Heading 2': { size: 12, bold: true, italic: true, color: '404040', before: 10, after: 2 } },
        { key: 'casual', name: 'Casual',
          Title: { size: 32, bold: false, align: 'left', after: 6, color: 'accent' },
          'Heading 1': { size: 18, bold: false, color: 'accent', before: 12, after: 6 },
          'Heading 2': { size: 14, bold: false, color: '595959', before: 6, after: 4 } },
        { key: 'centered', name: 'Centered',
          Title: { size: 28, caps: true, align: 'center', after: 6, color: 'accent' },
          'Heading 1': { size: 14, caps: true, bold: true, align: 'center', color: 'accent', before: 18, after: 6 },
          'Heading 2': { size: 12, bold: true, align: 'center', color: null, before: 10, after: 4 } },
        { key: 'lines', name: 'Lines (Simple)',
          Title: { size: 28, align: 'left', after: 6, color: null, border: 'accent' },
          'Heading 1': { size: 16, bold: false, color: 'accent', before: 18, after: 6, border: 'accent' },
          'Heading 2': { size: 13, bold: false, color: 'accent', before: 10, after: 4 } },
        { key: 'distinctive', name: 'Lines (Distinctive)',
          Title: { size: 30, bold: true, caps: true, align: 'left', after: 8, color: '262626', border: '262626' },
          'Heading 1': { size: 14, bold: true, caps: true, color: '262626', before: 18, after: 6, border: '262626' },
          'Heading 2': { size: 12, bold: true, color: 'accent', before: 10, after: 4 } },
        { key: 'shaded', name: 'Shaded',
          Title: { size: 28, bold: false, align: 'left', after: 6, color: 'FFFFFF', shading: 'accent' },
          'Heading 1': { size: 14, bold: true, color: 'FFFFFF', before: 18, after: 6, shading: 'accent' },
          'Heading 2': { size: 12, bold: true, color: 'accent', before: 10, after: 4, shading: 'E7E6E6' } },
        { key: 'minimalist', name: 'Minimalist',
          Title: { size: 24, bold: false, caps: true, align: 'left', after: 12, color: '595959' },
          'Heading 1': { size: 12, bold: true, caps: true, color: '404040', before: 18, after: 4 },
          'Heading 2': { size: 11, bold: true, color: '595959', before: 10, after: 2 } }
    ];

    return P;
});

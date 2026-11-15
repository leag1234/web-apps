/*
 * SPDX-FileCopyrightText: 2026 ElixirOffice contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 *
 * Common.UI.WordArtRibbon: the "WordArt Styles" group of the Shape Format
 * contextual tabs (Word and PowerPoint): a Quick Styles gallery (the Insert >
 * Text Art presets, applied to the text of the selected shape), Text Fill, Text
 * Outline and Text Effects. The engine has one text effect, Transform (warp), so
 * that is the only item of the menu. Everything goes through the same
 * asc_TextArtProperties calls as the Text Art pane of the right panel.
 *
 * The editors supply a small host object:
 *   host.add(slotId, component)  render the component into its slot of the tab, register it for locks
 *   host.lock                    lock list for the new controls
 *   host.apply(fn)               fn(asc_CShapeProperty) fills the properties, the host applies them
 *   host.api()                   the editor api
 *   host.collection()            Common.Collections.TextArt (the Text Art presets)
 *   host.fillTextArt()           fills that collection from the api (when still empty)
 *   host.strings                 txtWordArtQuickStyles, txtTextFill, txtTextOutline, txtTextEffects, txtTransform,
 *                                txtNoFill, txtNoOutline, txtWeight, txtNoTransform
 *   host.ids                     {styles, fill, outline, effects}: slot ids
 */
define([
    'core',
    'common/main/lib/component/Button',
    'common/main/lib/component/ColorButton',
    'common/main/lib/component/DataView',
    'common/main/lib/component/Menu'
], function () {
    'use strict';

    var WEIGHTS = [0.25, 0.5, 0.75, 1, 1.5, 2.25, 3, 4.5, 6];
    var uid = 0;

    Common.UI.WordArtRibbon = {

        WEIGHTS: WEIGHTS,

        // Builds the four controls; returns {controls: [...], buttons: [colour buttons], update: fn(textArtProps)}
        create: function (host) {
            var s = host.strings, ids = host.ids, lock = host.lock, n = ++uid, state = {weight: 1, fill: '000000', line: '000000'};

            var textArt = function (fill) {
                host.apply(function (shape) {
                    var props = new Asc.asc_TextArtProperties();
                    fill(props);
                    shape.put_TextArtProperties(props);
                });
            };

            // Quick Styles
            var btnStyles = new Common.UI.Button({
                cls: 'btn-toolbar x-huge icon-top', iconCls: 'toolbar__icon btn-elx-wordart-styles', caption: s.txtWordArtQuickStyles,
                lock: lock, dataHint: '1', dataHintDirection: 'bottom', dataHintOffset: 'small',
                menu: new Common.UI.Menu({style: 'min-width: 0;', items: [
                    {template: _.template('<div id="elx-wordart-styles-' + n + '" style="width: 330px; margin: 0 4px;"></div>')}
                ]})
            });
            host.add(ids.styles, btnStyles);
            var stylesBuilt = false;
            btnStyles.menu.on('show:before', function (menu) {
                if (stylesBuilt) return;
                var collection = host.collection();
                if (collection.length < 1) host.fillTextArt();
                if (collection.length < 1) return;
                stylesBuilt = true;
                var picker = new Common.UI.DataView({
                    el: $('#elx-wordart-styles-' + n, menu.$el),
                    store: collection,
                    parentMenu: menu,
                    outerMenu: {menu: menu, index: 0},
                    showLast: false,
                    restoreHeight: 300,
                    itemTemplate: _.template('<div class="item-art"><img src="<%= imageUrl %>" id="<%= id %>" style="width:50px;height:50px;"></div>')
                });
                picker.on('item:click', function (picker, item, record, e) {
                    record && textArt(function (p) { p.asc_putStyle(record.get('data')); });
                    if (e.type !== 'click') menu.hide();
                });
                menu.setInnerMenu([{menu: picker, index: 0}]);
            });

            // Text Fill
            var btnFill = new Common.UI.ButtonColored({
                cls: 'btn-toolbar', iconCls: 'toolbar__icon btn-elx-text-fill', caption: s.txtTextFill,
                split: true, menu: true, transparent: true, eyeDropper: true, colorLine: false, lock: lock,
                dataHint: '1', dataHintDirection: 'left', dataHintOffset: 'small'
            });
            host.add(ids.fill, btnFill);
            var applyFill = function (color) {
                textArt(function (p) {
                    var fill = new Asc.asc_CShapeFill();
                    if (color === 'transparent') {
                        fill.put_type(Asc.c_oAscFill.FILL_TYPE_NOFILL);
                        fill.put_fill(null);
                    } else {
                        fill.put_type(Asc.c_oAscFill.FILL_TYPE_SOLID);
                        fill.put_fill(new Asc.asc_CFillSolid());
                        fill.get_fill().put_color(Common.Utils.ThemeColor.getRgbColor(color));
                    }
                    p.asc_putFill(fill);
                });
            };
            btnFill.on('color:select', function (btn, color) { state.fill = color; applyFill(color); host.done && host.done(); });
            btnFill.on('click', function () { applyFill(btnFill.color || state.fill); host.done && host.done(); });

            // Text Outline: colour and weight
            var weight = new Common.UI.MenuItem({caption: s.txtWeight, menu: new Common.UI.Menu({menuAlign: 'tl-tr', items: _.map(WEIGHTS, function (v) {
                return {caption: v + ' pt', value: v, checkable: true, toggleGroup: 'elx-wordart-weight-' + n};
            })})});
            var btnOutline = new Common.UI.ButtonColored({
                cls: 'btn-toolbar', iconCls: 'toolbar__icon btn-elx-text-outline', caption: s.txtTextOutline,
                split: true, menu: true, transparent: true, eyeDropper: true, colorLine: false, lock: lock,
                additionalItemsAfter: [{caption: '--'}, weight],
                dataHint: '1', dataHintDirection: 'left', dataHintOffset: 'small'
            });
            host.add(ids.outline, btnOutline);
            var applyLine = function (color, pt) {
                textArt(function (p) {
                    var stroke = new Asc.asc_CStroke();
                    if (color === 'transparent') stroke.put_type(Asc.c_oAscStrokeType.STROKE_NONE);
                    else {
                        stroke.put_type(Asc.c_oAscStrokeType.STROKE_COLOR);
                        stroke.put_color(Common.Utils.ThemeColor.getRgbColor(color));
                        stroke.put_width((pt || state.weight) * 25.4 / 72);
                    }
                    p.asc_putLine(stroke);
                });
            };
            btnOutline.on('color:select', function (btn, color) { state.line = color; applyLine(color); host.done && host.done(); });
            btnOutline.on('click', function () { applyLine(btnOutline.color || state.line); host.done && host.done(); });
            weight.menu.on('item:click', function (menu, item) {
                state.weight = item.value;
                var c = btnOutline.color && btnOutline.color !== 'transparent' ? btnOutline.color : state.line;
                state.line = c;
                applyLine(c, item.value);
                host.done && host.done();
            });

            // Text Effects: Transform
            var tid = 'elx-wordart-transform-' + n;
            var transform = new Common.UI.MenuItem({caption: s.txtTransform, menu: new Common.UI.Menu({menuAlign: 'tl-tr', style: 'min-width: 0;', items: [
                {template: _.template('<div id="' + tid + '" style="width: 330px; margin: 0 4px;"></div>')}
            ]})});
            var btnEffects = new Common.UI.Button({
                cls: 'btn-toolbar', iconCls: 'toolbar__icon btn-elx-text-effects', caption: s.txtTextEffects,
                lock: lock, dataHint: '1', dataHintDirection: 'left', dataHintOffset: 'small',
                menu: new Common.UI.Menu({items: [transform]})
            });
            host.add(ids.effects, btnEffects);
            var transformBuilt = false;
            transform.menu.on('show:before', function (menu) {
                if (transformBuilt) return;
                var list = host.api() && host.api().asc_getPropertyEditorTextArts && host.api().asc_getPropertyEditorTextArts();
                if (!list || !list.length) return;
                transformBuilt = true;
                var store = new Common.UI.DataViewStore(_.map(list, function (it) {
                    return {imageUrl: it.Image, type: it.Type};
                }));
                var picker = new Common.UI.DataView({
                    el: $('#' + tid, menu.$el),
                    store: store,
                    parentMenu: menu,
                    outerMenu: {menu: menu, index: 0},
                    showLast: false,
                    restoreHeight: 300,
                    itemTemplate: _.template('<div class="item-art"><img src="<%= imageUrl %>" id="<%= id %>" style="width:50px;height:50px;"></div>')
                });
                picker.on('item:click', function (picker, item, record, e) {
                    record && textArt(function (p) { p.asc_putForm(record.get('type')); });
                    host.done && host.done();
                    if (e.type !== 'click') btnEffects.menu.hide();
                });
                menu.setInnerMenu([{menu: picker, index: 0}]);
            });

            var buttons = [btnFill, btnOutline];
            return {
                btnStyles: btnStyles, btnFill: btnFill, btnOutline: btnOutline, btnEffects: btnEffects, weight: weight,
                controls: [btnStyles, btnFill, btnOutline, btnEffects],
                buttons: buttons,
                // theme colours changed: refresh the two colour pickers
                updateColors: function (effect, standard) {
                    _.each(buttons, function (b) {
                        if (!b.menu || !_.isObject(b.menu)) { b.setMenu(); }
                        b.getPicker().updateColors(effect, standard);
                    });
                    btnFill.color === undefined && btnFill.setColor && btnFill.setColor(state.fill);
                    btnOutline.color === undefined && btnOutline.setColor && btnOutline.setColor(state.line);
                }
            };
        }
    };
});

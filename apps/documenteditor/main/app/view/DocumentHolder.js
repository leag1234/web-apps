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
 *  DocumentHolder.js
 *
 *  DocumentHolder view
 *
 *  Created on 1/11/14
 *
 */

define([
    'jquery',
    'underscore',
    'backbone',
    'gateway',
    'common/main/lib/util/utils',
    'common/main/lib/component/Menu',
    'common/main/lib/component/Calendar'
], function ($, _, Backbone, gateway) { 'use strict';

    DE.Views.DocumentHolder =  Backbone.View.extend(_.extend({
        el: '#editor_sdk',

        // Compile our stats template
        template: null,

        // Delegated events for creating new items, and clearing completed ones.
        events: {
        },

        initialize: function () {
            this._currentMathObj = undefined;
            this._currentSpellObj = undefined;
            this._currentParaObjDisabled = false;
            this._currentTranslateObj = this;
            this._currLang        = {};
            this._isDisabled = false;
            this._preventCustomClick = null;
            this._hasCustomItems = false;
            this._docProtection = {
                isReadOnly: false,
                isReviewOnly: false,
                isFormsOnly: false,
                isCommentsOnly: false
            };
            this._unprotectedRegion = {
                canEditText: false,
                canEditPara: false,
                canInsObject: false
            };
            this._langs = null;
        },

        render: function () {
            this.fireEvent('render:before', this);

            this.cmpEl = $(this.el);

            this.fireEvent('render:after', this);
            return this;
        },

        setApi: function(o) {
            this.api = o;
            return this;
        },

        createDelayedElements: function() {},

        createDelayedElementsViewer: function() {},

        createDelayedElementsPDFViewer: function() {},

        createDelayedElementsPDFForms: function() {},

        setMode: function(m) {
            this.mode = m;
            this._fillFormMode = !this.mode.isEdit && this.mode.canFillForms;
            return this;
        },

        addWordVariants: function(isParagraph) {
            var me = this;
            if (!me.textMenu || !me.textMenu.isVisible() && !me.tableMenu.isVisible()) return;

            if (_.isUndefined(isParagraph)) {
                isParagraph = me.textMenu.isVisible();
            }

            me.clearWordVariants(isParagraph);

            var moreMenu  = (isParagraph) ? me.menuSpellMorePara : me.menuSpellMoreTable;
            var spellMenu = (isParagraph) ? me.menuSpellPara : me.menuSpellTable;
            var menu = (isParagraph) ? me.textMenu : me.tableMenu;
            var variants = me._currentSpellObj.get_Variants();

            // as in Word: the best suggestions at the top of the menu (no sub-menu for the others), the menu has no
            // "No Spelling Suggestions" entry when there are some
            moreMenu.setVisible(false);
            if (variants.length > 0) {
                _.each(variants.slice(0, me.maxSpellVariants), function(variant, index) {
                    var mnu = new Common.UI.MenuItem({
                        caption     : variant,
                        spellword   : true,
                        disabled    : me._currentParaObjDisabled
                    }).on('click', function(item, e) {
                        if (me.api) {
                            me.api.asc_replaceMisspelledWord(item.caption, me._currentSpellObj);
                            me.fireEvent('editcomplete', me);
                        }
                    });
                    menu.insertItem(index, mnu);
                });
                spellMenu.setVisible(false);
            } else {
                spellMenu.setVisible(true);
                spellMenu.setCaption(me.noSpellVariantsText);
                spellMenu.setDisabled(true);
            }
        },

        maxSpellVariants: 5,

        clearWordVariants: function(isParagraph) {
            var me = this;
            var spellMenu = (isParagraph) ? me.textMenu : me.tableMenu;

            for (var i = 0; i < spellMenu.items.length; i++) {
                if (spellMenu.items[i].options.spellword) {
                    if (spellMenu.checkeditem == spellMenu.items[i]) {
                        spellMenu.checkeditem = undefined;
                        spellMenu.activeItem  = undefined;
                    }

                    spellMenu.removeItem(spellMenu.items[i]);
                    i--;
                }
            }
            (isParagraph) ? me.menuSpellMorePara.menu.removeAll() : me.menuSpellMoreTable.menu.removeAll();

            me.menuSpellMorePara.menu.checkeditem   = undefined;
            me.menuSpellMorePara.menu.activeItem    = undefined;
            me.menuSpellMoreTable.menu.checkeditem  = undefined;
            me.menuSpellMoreTable.menu.activeItem   = undefined;
        },

        // the Language sub-menus of the spelling menus are filled when they open (language at the cursor, recently
        // used ones, "Set Proofing Language..."): the list of languages is only kept here
        setLanguages: function(langs){
            if (langs && langs.length > 0) this._langs = langs;
        },

        fillLanguageMenu: function(menu) {
            Common.Utils.fillProofingLanguageMenu(menu, this._langs || [], this._currLang.id, this.textSetProofingLanguage);
        },

        changeLanguageMenu: function(menu) {},

        textSetProofingLanguage: 'Set Proofing Language...',

        getControlLabel: function(props) {
            var type = props ? props.get_SpecificType() : Asc.c_oAscContentControlSpecificType.None;
            switch (type) {
                case Asc.c_oAscContentControlSpecificType.CheckBox:
                    var specProps = props.get_CheckBoxPr();
                    return (typeof specProps.get_GroupKey() !== 'string') ? this.textRemCheckBox : this.textRemRadioBox;
                case Asc.c_oAscContentControlSpecificType.ComboBox:
                    return this.textRemComboBox;
                case Asc.c_oAscContentControlSpecificType.DropDownList:
                    return this.textRemDropdown;
                case Asc.c_oAscContentControlSpecificType.Picture:
                case Asc.c_oAscContentControlSpecificType.Signature:
                    return this.textRemPicture;
                default:
                    return this.textRemField;
            }
        },

        createEquationMenu: function(toggleGroup, menuAlign) {
            return new Common.UI.Menu({
                cls: 'ppm-toolbar shifted-right',
                menuAlign: menuAlign,
                items   : [
                    new Common.UI.MenuItem({
                        caption     : this.currProfText,
                        iconCls     : 'menu__icon btn-professional-equation',
                        type        : 'view',
                        value       : {all: false, linear: false}
                    }),
                    new Common.UI.MenuItem({
                        caption     : this.currLinearText,
                        iconCls     : 'menu__icon btn-linear-equation',
                        type        : 'view',
                        value       : {all: false, linear: true}
                    }),
                    new Common.UI.MenuItem({
                        caption     : this.allProfText,
                        iconCls     : 'menu__icon btn-professional-equation',
                        type        : 'view',
                        value       : {all: true, linear: false}
                    }),
                    new Common.UI.MenuItem({
                        caption     : this.allLinearText,
                        iconCls     : 'menu__icon btn-linear-equation',
                        type        : 'view',
                        value       : {all: true, linear: true}
                    }),
                    { caption     : '--' },
                    new Common.UI.MenuItem({
                        caption     : this.unicodeText,
                        checkable   : true,
                        checked     : false,
                        toggleGroup : toggleGroup,
                        type        : 'input',
                        value       : Asc.c_oAscMathInputType.Unicode
                    }),
                    new Common.UI.MenuItem({
                        caption     : this.latexText,
                        checkable   : true,
                        checked     : false,
                        toggleGroup : toggleGroup,
                        type        : 'input',
                        value       : Asc.c_oAscMathInputType.LaTeX
                    }),
                    { caption     : '--' },
                    new Common.UI.MenuItem({
                        caption     : this.eqToInlineText,
                        isEquationInline: false,
                        type        : 'mode'
                    }),
                    new Common.UI.MenuItem({
                        caption     : this.hideEqToolbar,
                        isToolbarHide: false,
                        type        : 'hide',
                    })
                ]
            });
        },

        focus: function() {
            var me = this;
            _.defer(function(){  me.cmpEl.focus(); }, 50);
        },

        SetDisabled: function(state, canProtect, fillFormMode) {
            this._isDisabled = state;
            this._canProtect =  state ? canProtect : true;
            this._fillFormMode = state ? fillFormMode : false;
        },

        addEquationMenu: function() {},

        clearEquationMenu: function() {},

        equationCallback: function() {},

        initEquationMenu: function() {},

        updateCustomItems: function() {},

        clearCustomItems: function() {},

        parseIcons: function() {}

}, DE.Views.DocumentHolder || {}));
});

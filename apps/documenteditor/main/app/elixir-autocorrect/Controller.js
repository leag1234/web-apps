/*
 * SPDX-FileCopyrightText: 2026 ElixirOffice contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 *
 * ElixirOffice AutoCorrect (document editor): Word's "Replace text as you type". Common typos
 * per language are corrected when a space, a punctuation mark or Enter follows the word
 * (sdkjs CRunAutoCorrect, AUTOCORRECT_FLAGS_TEXT_REPLACE); Ctrl+Z right after a correction
 * restores what was typed. (c), (r), (tm), straight-to-curly quotes and the capital at the start
 * of a sentence are the engine's own autocorrections, on by default.
 * The list is edited in File > Options > Proofing > AutoCorrect Options, "Replace Text" page
 * (Common.Views.AutoCorrectDialog calls ElixirTextAutoCorrect.renderPanel); the user's entries
 * stay in the browser (localStorage). The only hooks are the controller and module
 * declarations in app.js / app_dev.js and the panel in AutoCorrectDialog.
 */
define([
    'core'
], function () {
    'use strict';

    var KEY_ON = 'de-elixir-autocorrect-replace', KEY_USER = 'de-elixir-autocorrect-user';

    // Word's most common entries for each language (lower-case typed word -> correction)
    var DEFAULTS = {
        en: {
            'i': 'I', 'teh': 'the', 'adn': 'and', 'hte': 'the', 'taht': 'that', 'thier': 'their', 'recieve': 'receive', 'recieved': 'received',
            'wich': 'which', 'becuase': 'because', 'beleive': 'believe', 'definately': 'definitely', 'seperate': 'separate',
            'occured': 'occurred', 'occurence': 'occurrence', 'untill': 'until', 'tommorow': 'tomorrow', 'accomodate': 'accommodate',
            'acheive': 'achieve', 'alot': 'a lot', 'arguement': 'argument', 'begining': 'beginning', 'calender': 'calendar',
            'collegue': 'colleague', 'comming': 'coming', 'commited': 'committed', 'concious': 'conscious', 'embarass': 'embarrass',
            'enviroment': 'environment', 'existance': 'existence', 'foriegn': 'foreign', 'freind': 'friend', 'goverment': 'government',
            'happend': 'happened', 'independant': 'independent', 'knowlege': 'knowledge', 'neccessary': 'necessary',
            'noticable': 'noticeable', 'peice': 'piece', 'posession': 'possession', 'prefered': 'preferred', 'publically': 'publicly',
            'realy': 'really', 'refered': 'referred', 'relevent': 'relevant', 'succesful': 'successful', 'suprise': 'surprise',
            'truely': 'truly', 'wierd': 'weird', 'dont': 'don’t', 'doesnt': 'doesn’t', 'didnt': 'didn’t',
            'isnt': 'isn’t', 'wasnt': 'wasn’t', 'couldnt': 'couldn’t', 'shouldnt': 'shouldn’t', 'wouldnt': 'wouldn’t'
        },
        fr: {
            'acceuil': 'accueil', 'acceuillir': 'accueillir', 'apercu': 'aperçu', 'aujourdhui': 'aujourd’hui', 'biensur': 'bien sûr',
            'parmis': 'parmi', 'malgres': 'malgré', 'dailleurs': 'd’ailleurs', 'quelquechose': 'quelque chose',
            'developpement': 'développement', 'developper': 'développer', 'etre': 'être', 'deja': 'déjà', 'voila': 'voilà',
            'tres': 'très', 'apres': 'après', 'probleme': 'problème', 'systeme': 'système', 'interressant': 'intéressant',
            'addresse': 'adresse', 'coeur': 'cœur', 'oeuvre': 'œuvre', 'soeur': 'sœur', 'noeud': 'nœud',
            'oeil': 'œil', 'oeuf': 'œuf', 'rapellez': 'rappelez'
        },
        de: {
            'nciht': 'nicht', 'ihc': 'ich', 'udn': 'und', 'vieleicht': 'vielleicht', 'wiederspiegeln': 'widerspiegeln', 'seperat': 'separat',
            'standart': 'Standard', 'addresse': 'Adresse', 'agressiv': 'aggressiv', 'nähmlich': 'nämlich', 'wiedersprechen': 'widersprechen',
            'rythmus': 'Rhythmus', 'tollerant': 'tolerant', 'entg\u00fcltig': 'endgültig', 'poblem': 'Problem', 'gehts': 'geht’s'
        },
        es: {
            'atravez': 'a través', 'derrepente': 'de repente', 'enserio': 'en serio', 'tambien': 'también', 'despues': 'después',
            'asi': 'así', 'aqui': 'aquí', 'ademas': 'además', 'facil': 'fácil', 'dificil': 'difícil',
            'informacion': 'información', 'atencion': 'atención', 'tambein': 'también'
        },
        it: {
            'perche': 'perché', 'poiche': 'poiché', 'piu': 'più', 'puo': 'può', 'gia': 'già', 'cioe': 'cioè',
            'citta': 'città', 'propio': 'proprio', 'daccordo': 'd’accordo'
        }
    };

    var A = window.ElixirTextAutoCorrect = {
        defaults: DEFAULTS,

        enabled: function () { return Common.localStorage.getBool(KEY_ON, true); },

        user: function () {
            try { return JSON.parse(Common.localStorage.getItem(KEY_USER) || '{}') || {}; } catch (e) { return {}; }
        },

        saveUser: function (u) { Common.localStorage.setItem(KEY_USER, JSON.stringify(u)); },

        // the list used by the engine: defaults minus removed entries plus the user's own
        lists: function () {
            var u = this.user(), res = {};
            _.each(_.union(_.keys(DEFAULTS), _.keys(u)), function (lang) {
                var l = _.extend({}, DEFAULTS[lang] || {}), mine = u[lang] || {};
                _.each(mine.rem || [], function (k) { delete l[k]; });
                _.extend(l, mine.add || {});
                res[lang] = l;
            });
            return res;
        },

        apply: function (api) {
            api = api || this.api;
            api && api.asc_ElixirSetTextAutoCorrect && api.asc_ElixirSetTextAutoCorrect(this.enabled(), this.lists());
        },

        langOf: function (api) {
            var code = api && api.asc_getDefaultLanguage ? api.asc_getDefaultLanguage() : null,
                info = code && Common.util.LanguageInfo ? Common.util.LanguageInfo.getLocalLanguageName(code) : null,
                name = info && info[0];
            name = (name || 'en').split('-')[0].toLowerCase();
            return DEFAULTS[name] ? name : 'en';
        },

        // "Replace Text" page of the AutoCorrect dialog
        renderPanel: function ($el, dlg) {
            var me = this, api = dlg.api || this.api, s = dlg;
            var lang = this.langOf(api);
            var names = {en: 'English', fr: 'Français', de: 'Deutsch', es: 'Español', it: 'Italiano'};
            $el.html(
                '<div class="elx-ac-panel">' +
                '<div id="elx-ac-on" style="margin-bottom: 10px;"></div>' +
                '<div class="elx-ac-row"><label>' + s.textElxLanguage + '</label><select class="form-control elx-ac-lang">' +
                _.map(_.keys(DEFAULTS), function (l) { return '<option value="' + l + '"' + (l === lang ? ' selected' : '') + '>' + names[l] + '</option>'; }).join('') + '</select></div>' +
                '<div class="elx-ac-row"><div><label>' + s.textElxReplace + '</label><input type="text" class="form-control elx-ac-from" spellcheck="false"></div>' +
                '<div><label>' + s.textElxWith + '</label><input type="text" class="form-control elx-ac-to" spellcheck="false"></div></div>' +
                '<div class="elx-ac-list" role="listbox"></div>' +
                '<div class="elx-ac-buttons"><button type="button" class="btn normal dlg-btn elx-ac-add">' + s.textElxAdd + '</button>' +
                '<button type="button" class="btn normal dlg-btn elx-ac-del">' + s.textElxDelete + '</button>' +
                '<button type="button" class="btn normal dlg-btn elx-ac-reset">' + s.textElxReset + '</button></div></div>');
            var chk = new Common.UI.CheckBox({el: $el.find('#elx-ac-on'), labelText: s.textElxReplaceAsYouType, value: this.enabled()});
            chk.on('change', function (field, value) { Common.localStorage.setBool(KEY_ON, value === 'checked' || value === true); me.apply(api); });
            var $from = $el.find('.elx-ac-from'), $to = $el.find('.elx-ac-to'), $list = $el.find('.elx-ac-list'), $lang = $el.find('.elx-ac-lang');
            var fill = function () {
                var l = me.lists()[$lang.val()] || {}, q = ($from.val() || '').toLowerCase(), keys = _.keys(l).sort();
                $list.html(_.map(keys, function (k) {
                    return '<div class="elx-ac-item' + (k === q ? ' selected' : '') + '" role="option" data-key="' + Common.Utils.String.htmlEncode(k) + '"><span>' +
                        Common.Utils.String.htmlEncode(k) + '</span><span>' + Common.Utils.String.htmlEncode(l[k]) + '</span></div>';
                }).join(''));
                var $sel = $list.find('.selected');
                $sel.length && $sel[0].scrollIntoView({block: 'nearest'});
            };
            var edit = function (fn) {
                var u = me.user(), l = $lang.val(); u[l] = u[l] || {add: {}, rem: []}; u[l].add = u[l].add || {}; u[l].rem = u[l].rem || [];
                fn(u[l]); me.saveUser(u); me.apply(api); fill();
            };
            $el.find('input').on('focus', function () { api && api.asc_enableKeyEvents(false); });
            $from.on('input', fill);
            $lang.on('change', fill);
            $list.on('click', '.elx-ac-item', function () {
                var k = $(this).attr('data-key'), l = me.lists()[$lang.val()] || {};
                $from.val(k); $to.val(l[k] || ''); fill();
            });
            $el.find('.elx-ac-add').on('click', function () {
                var k = ($from.val() || '').trim().toLowerCase(), v = ($to.val() || '').trim();
                if (!k || !v || k === v.toLowerCase() && k === v) return;
                edit(function (e) { e.add[k] = v; e.rem = _.without(e.rem, k); });
            });
            $el.find('.elx-ac-del').on('click', function () {
                var k = ($from.val() || '').trim().toLowerCase();
                if (!k) return;
                edit(function (e) { delete e.add[k]; DEFAULTS[$lang.val()] && DEFAULTS[$lang.val()][k] && e.rem.indexOf(k) < 0 && e.rem.push(k); });
                $from.val(''); $to.val(''); fill();
            });
            $el.find('.elx-ac-reset').on('click', function () {
                var u = me.user(); delete u[$lang.val()]; me.saveUser(u); me.apply(api); fill();
            });
            fill();
        }
    };

    DE.Controllers.ElixirAutoCorrect = Backbone.Controller.extend({
        initialize: function () {
            var me = this;
            Common.NotificationCenter.on('app:ready', function (mode) { _.defer(function () { me.onAppReady(mode); }); });
        },
        onLaunch: function () {},
        onAppReady: function (mode) {
            var main = DE.getController('Main');
            if (!mode || !mode.isEdit || !main || !main.api) return;
            A.api = main.api;
            A.apply(main.api);
        }
    });
});

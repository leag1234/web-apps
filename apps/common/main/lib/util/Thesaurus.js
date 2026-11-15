/*
 * SPDX-FileCopyrightText: 2026 ElixirOffice contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 *
 * Common.util.Thesaurus: the offline MyThes thesaurus of the dictionaries folder (17 languages,
 * see docs/thesaurus.md). A thesaurus is loaded on first use (its index only), the meanings of a
 * word are read from the .dat file with an HTTP range request; nothing leaves the server.
 *
 *   files(lcid)          the candidate file names for a Windows language id
 *   load(lcid)           Promise of {dat, enc, index} or null when there is none for the language
 *   lookup(thes, word)   Promise of [{pos, words: []}] or null
 *   parseMeanings(lines, key)   the .dat lines of a word to meanings (exposed for tests)
 *
 * MyThes files: <dir>.idx is "word|offset" per line, <dir>.dat holds "word|n" then n lines "(pos)|syn|syn...".
 */

define([
    'core'
], function () {
    'use strict';

    Common.util = Common.util || {};

    Common.util.Thesaurus = new (function () {
        var cache = {};

        // the thesaurus of the language of the text, else the one of its main variant (de-AT, de-CH: de-DE)
        this.DEFAULT = {ar: 'ar', ca: 'ca_ES', da: 'da_DK', de: 'de_DE', en: 'en_US', es: 'es_ES', fr: 'fr_FR', gl: 'gl_ES',
            it: 'it_IT', pl: 'pl_PL', ro: 'ro_RO', ru: 'ru_RU', sk: 'sk_SK', sl: 'sl_SI', sv: 'sv_SE', uk: 'uk_UA'};

        this.base = '../../../../dictionaries/';

        this.files = function (lcid) {
            var info = Common.util.LanguageInfo.getLocalLanguageName(lcid), code = info && info[0] ? info[0].replace(/-/g, '_') : 'en_US',
                dirs = [code], main = this.DEFAULT[code.split('_')[0].toLowerCase()];
            main && main !== code && dirs.push(main);
            var list = [];
            _.each(dirs, function (d) { list.push(d + '/' + d + '_thes', d + '/th_' + d + '_v2', d + '/th_' + d); });
            return list;
        };

        this.load = function (lcid) {
            var base = this.base, files = this.files(lcid), key = String(lcid);
            if (cache[key]) return cache[key];
            var tryNext = function (i) {
                if (i >= files.length) return Promise.resolve(null);
                return fetch(base + files[i] + '.idx').then(function (r) {
                    if (!r.ok) throw new Error('missing');
                    return r.arrayBuffer();
                }).then(function (buf) {
                    var head = new TextDecoder('latin1').decode(new Uint8Array(buf, 0, Math.min(40, buf.byteLength))).split(/\r?\n/)[0].replace(/^﻿|^ï»¿/, '').trim();
                    var enc = /utf-?8/i.test(head) ? 'utf-8' : 'latin1';
                    var lines = new TextDecoder(enc).decode(buf).split(/\r?\n/), index = new Map();
                    for (var n = 2; n < lines.length; n++) {
                        var bar = lines[n].lastIndexOf('|');
                        bar > 0 && index.set(lines[n].substring(0, bar).toLowerCase(), +lines[n].substring(bar + 1));
                    }
                    return {dat: base + files[i] + '.dat', enc: enc, index: index};
                }).catch(function () { return tryNext(i + 1); });
            };
            cache[key] = tryNext(0);
            return cache[key];
        };

        this.parseMeanings = function (lines, key) {
            var start = 0;
            while (start < lines.length - 1 && lines[start].toLowerCase().indexOf(key + '|') !== 0) start++;
            lines = lines.slice(start);
            var first = (lines[0] || '').split('|'), count = +first[1] || 0, meanings = [];
            for (var i = 1; i <= count && i < lines.length; i++) {
                var parts = lines[i].split('|');
                meanings.push({pos: (parts[0] || '').replace(/[()]/g, ''), words: parts.slice(1).filter(Boolean)});
            }
            return meanings;
        };

        this.lookup = function (thes, word) {
            var key = String(word || '').toLowerCase(), offset = thes && thes.index.get(key), me = this;
            if (offset === undefined) return Promise.resolve(null);
            // the offsets of some files do not count the byte order mark of the .dat: read a little before
            var from = Math.max(0, offset - 8);
            return fetch(thes.dat, {headers: {Range: 'bytes=' + from + '-' + (offset + 16383)}}).then(function (r) { return r.arrayBuffer(); }).then(function (buf) {
                return me.parseMeanings(new TextDecoder(thes.enc).decode(buf).split(/\r?\n/), key);
            });
        };
    })();

    return Common.util.Thesaurus;
});

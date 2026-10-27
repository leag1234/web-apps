/*
 * SPDX-FileCopyrightText: 2026 ElixirOffice contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 *
 * Elixir local fonts, engine side: use the fonts installed on the user's computer in the editor
 * (Local Font Access API: window.queryLocalFonts(), Chromium 103+, secure context, user consent).
 *
 * Design (see docs/localfonts.md):
 *  - the editor engine (sdkjs) measures and draws with the font files listed in its runtime font index
 *    (AscFonts.g_font_files / g_font_infos / g_map_font_index), loaded through a CFontFileLoader. We add
 *    entries to that index, one per local family, whose loader is lazy: the font file (FontData.blob()) is
 *    read only when the document or the user needs the face, then handed to the engine from memory
 *    (CFontFileLoader.LoadFontFromData). Nothing is uploaded: the bytes stay in this page.
 *  - a local font is registered only under a name the server does not have. The engine then asks for it by
 *    its exact name, which wins over the substitute it would have picked (same rule as docs/fonts.md:
 *    a real installed font beats a metric clone). Two thin wrappers make that choice (GetFontFileWeb,
 *    GetFontFile) and one wrapper on the loader (LoadDocumentFonts) registers the fonts before the
 *    document fonts are loaded, so a consenting user gets the right layout from the first paint.
 *  - without consent, nothing is read: the list of fonts of this computer is not even requested.
 */

define([], function () {
    'use strict';

    var LF = (window.Common = window.Common || {}).ElixirLocalFonts = window.Common.ElixirLocalFonts || {};

    var KEY_CONSENT = 'elixir.localfonts.consent';

    // Fonts the server replaces by a metric-identical clone, and the symbol fonts the engine always asks for
    // (a document that merely has default styles): not worth an invitation, or a "not found" notice, on their own.
    // They are still taken from this computer once the user agreed.
    LF.COMMON_FONTS = ['Arial', 'Times New Roman', 'Courier New', 'Symbol', 'Wingdings'];

    // ---- storage ---------------------------------------------------------------------------------
    function lsGet(k) { try { return window.localStorage.getItem(k); } catch (e) { return null; } }
    function lsSet(k, v) { try { v === null ? window.localStorage.removeItem(k) : window.localStorage.setItem(k, v); } catch (e) { /* private mode */ } }

    LF.hasConsent = function () { return lsGet(KEY_CONSENT) === '1'; };
    LF.setConsent = function (on) { lsSet(KEY_CONSENT, on ? '1' : null); };

    // The preload bridge belongs to the same-origin desktop start window.
    LF.desktopBridge = function () {
        try {
            var bridge = window.eoDesktop || window.parent && window.parent.eoDesktop;
            return bridge && typeof bridge.fontConsent === 'function' ? bridge : null;
        } catch (e) { return null; }
    };

    // ---- support -----------------------------------------------------------------------------------
    // 'ok' | 'insecure' (not HTTPS) | 'unsupported' (Firefox, Safari...)
    LF.support = function () {
        if (LF.desktopBridge()) return 'ok';
        if (!window.isSecureContext) return 'insecure';
        if (typeof window.queryLocalFonts !== 'function') return 'unsupported';
        return 'ok';
    };

    // 'granted' | 'prompt' | 'denied' | null (cannot be asked)
    LF.permission = function () {
        var bridge = LF.desktopBridge();
        if (bridge) return typeof bridge.fontPermission === 'function' ?
            bridge.fontPermission().then(function (allowed) { return allowed ? 'granted' : 'prompt'; }) : Promise.resolve('prompt');
        try {
            return navigator.permissions.query({ name: 'local-fonts' }).then(function (s) { return s.state; }, function () { return null; });
        } catch (e) { return Promise.resolve(null); }
    };

    // ---- state shared with the UI ------------------------------------------------------------------
    LF.docFonts = [];          // font names requested by the open document
    LF.trace = [];             // last steps (for diagnostics and tests)
    var trace = function (m) { LF.trace.push(Math.round(performance.now()) + ' ' + m); if (LF.trace.length > 60) LF.trace.shift(); };
    LF.status = 'off';         // off | loading | on | denied | blocked | error
    LF.count = 0;              // fonts registered from this computer
    LF.listeners = [];
    LF.onChange = function (f) { LF.listeners.push(f); };
    function changed() { LF.listeners.forEach(function (f) { try { f(LF); } catch (e) { /* ui */ } }); }
    LF.setStatus = function (s) { LF.status = s; changed(); };

    // ---- engine ------------------------------------------------------------------------------------
    function engine() {
        var A = window.AscFonts;
        if (!A || !A.g_font_infos || !A.g_font_files || !A.g_map_font_index || !A.g_fontApplication || !A.CFontFileLoader || !A.CFontInfo) return null;
        if (!window.AscCommon || !AscCommon.g_font_loader) return null;
        return A;
    }
    LF.engineReady = function () { return !!engine(); };

    var norm = function (s) { return String(s || '').toLowerCase().replace(/[^a-z0-9À-ɏͰ-῿぀-鿿]/g, ''); };
    var local = { byName: {}, byNorm: {} };    // registered local fonts: name -> {name, key}
    LF.isLocal = function (name) { return !!(name && (local.byName[name] || local.byNorm[norm(name)])); };
    LF.resolveName = function (name) {
        var e = local.byName[name] || local.byNorm[norm(name)];
        return e ? e.name : null;
    };
    LF.localNames = function () { return Object.keys(local.byName); };

    // Names of the document fonts that the server does not have and this computer does not provide (yet)
    LF.neededNames = function () {
        var A = engine();
        var seen = {};
        return (LF.docFonts || []).filter(function (n) {
            if (!n || seen[n] || /^Embedded: /.test(n) || n === 'ASCW3') return false;
            seen[n] = true;
            if (LF.COMMON_FONTS.indexOf(n) >= 0) return false;
            if (LF.isLocal(n)) return false;
            return !A || A.g_map_font_index[n] === undefined;
        });
    };

    // ---- sfnt reading (face index of a collection, named instance of a variable font) -----------------
    // Reads the name table (full name, PostScript name) of the font at offset "off", and the named instances of its
    // 'fvar' table (variable font): [{name, ps}] in the order FreeType numbers them.
    function sfntInfo(d, off) {
        var u16 = function (o) { return (d[o] << 8) | d[o + 1]; };
        var u32 = function (o) { return ((d[o] << 24) | (d[o + 1] << 16) | (d[o + 2] << 8) | d[o + 3]) >>> 0; };
        var res = { full: '', ps: '', variable: false, instances: [] };
        var n = u16(off + 4), nameOff = -1, fvarOff = -1;
        for (var i = 0; i < n; i++) {
            var r = off + 12 + 16 * i;
            var tag = String.fromCharCode(d[r], d[r + 1], d[r + 2], d[r + 3]);
            if (tag === 'name') nameOff = u32(r + 8);
            else if (tag === 'fvar') fvarOff = u32(r + 8);
        }
        var count = 0, store = 0;
        if (nameOff >= 0 && nameOff + 6 <= d.length) { count = u16(nameOff + 2); store = nameOff + u16(nameOff + 4); }
        var nameOf = function (want) {
            var best = '';
            for (var k = 0; k < count; k++) {
                var o = nameOff + 6 + 12 * k;
                var plat = u16(o), id = u16(o + 6), len = u16(o + 8), so = store + u16(o + 10);
                if (id !== want || so + len > d.length) continue;
                var s = '';
                if (plat === 3 || plat === 0) for (var j = 0; j + 1 < len; j += 2) s += String.fromCharCode(u16(so + j));
                else for (var m = 0; m < len; m++) s += String.fromCharCode(d[so + m]);
                if (s && (!best || plat === 3)) best = s;
            }
            return best;
        };
        res.full = nameOf(4);
        res.ps = nameOf(6);
        if (fvarOff >= 0 && fvarOff + 16 <= d.length) {
            res.variable = true;
            var axesOff = fvarOff + u16(fvarOff + 4), axisCount = u16(fvarOff + 8), axisSize = u16(fvarOff + 10);
            var instCount = u16(fvarOff + 12), instSize = u16(fvarOff + 14);
            for (var q = 0; q < instCount && q < 512; q++) {
                var io = axesOff + axisCount * axisSize + q * instSize;
                if (io + 4 > d.length) break;
                var psId = (instSize >= 6 + 4 * axisCount) ? u16(io + 4 + 4 * axisCount) : 0;
                res.instances.push({ name: nameOf(u16(io)), ps: psId && psId !== 0xFFFF ? nameOf(psId) : '' });
            }
        }
        return res;
    }
    LF._sfntInfo = sfntInfo;

    // Face index to give the engine (FreeType) for "fd" in the file "d": the index of the face in a collection,
    // plus, for a variable font, the named instance (instance number + 1 in the upper 16 bits) that "fd" is.
    function faceIndexOf(d, fd) {
        var ps = String(fd.postscriptName || '').toLowerCase(), full = String(fd.fullName || '').toLowerCase();
        var style = String(fd.style || '').toLowerCase();
        var idx = 0, off = 0;
        try {
            if (d.length >= 16 && String.fromCharCode(d[0], d[1], d[2], d[3]) === 'ttcf') {
                var n = ((d[8] << 24) | (d[9] << 16) | (d[10] << 8) | d[11]) >>> 0;
                for (var i = 0; i < n && i < 64; i++) {
                    var o = 12 + 4 * i, fo = ((d[o] << 24) | (d[o + 1] << 16) | (d[o + 2] << 8) | d[o + 3]) >>> 0;
                    var inf = sfntInfo(d, fo);
                    if ((ps && inf.ps.toLowerCase() === ps) || (full && inf.full.toLowerCase() === full)) { idx = i; off = fo; break; }
                }
            }
            var info = sfntInfo(d, off);
            if (info.variable) {
                for (var k = 0; k < info.instances.length; k++) {
                    var it = info.instances[k];
                    if ((ps && it.ps.toLowerCase() === ps) || (it.name && it.name.toLowerCase() === style)) { idx |= (k + 1) << 16; break; }
                }
            }
        } catch (e) { /* damaged tables: face 0 */ }
        return idx;
    }

    // ---- face selection ------------------------------------------------------------------------------
    var REGULAR_WORDS = { regular: 1, normal: 1, roman: 1, standard: 1, text: 1, plain: 1 };
    var ALT_REGULAR = { book: 1, medium: 1 };
    function classify(style) {
        var toks = String(style || '').toLowerCase().split(/[\s_\-]+/).filter(Boolean);
        var italic = false, bold = false, other = 0, alt = 0;
        toks.forEach(function (t) {
            if (t === 'italic' || t === 'oblique' || t === 'slanted' || t === 'inclined') italic = true;
            else if (t === 'bold') bold = true;
            else if (REGULAR_WORDS[t]) { /* nothing */ }
            else if (ALT_REGULAR[t]) alt++;
            else other++;
        });
        return { italic: italic, bold: bold, plain: other === 0, alt: alt, other: other };
    }
    // lower is better; Infinity: not usable for the slot
    function score(c, wantBold, wantItalic) {
        if (c.bold !== wantBold || c.italic !== wantItalic) return Infinity;
        if (c.other) return 10 + c.other;
        return c.alt ? 1 : 0;
    }
    function pickSlots(faces) {
        var cls = faces.map(function (fd) { return { fd: fd, c: classify(fd.style) }; });
        var best = function (b, i) {
            var r = null, rs = Infinity;
            cls.forEach(function (x) { var s = score(x.c, b, i); if (s < rs) { rs = s; r = x.fd; } });
            // a weight/width variant is not the bold or italic of the family: only plain styles, except for regular
            return (rs <= 1 || (!b && !i && rs < Infinity)) ? r : null;
        };
        var res = { R: best(false, false), I: best(false, true), B: best(true, false), BI: best(true, true) };
        if (!res.R) {   // only variants (e.g. "Light" alone): the first face stands for the regular style
            var any = cls.filter(function (x) { return !x.c.bold && !x.c.italic; })[0] || cls[0];
            res.R = any ? any.fd : null;
        }
        return res;
    }

    // ---- registration --------------------------------------------------------------------------------
    var SLOTS = ['R', 'I', 'B', 'BI'];

    function makeLoader(A, id, fd) {
        var L = new A.CFontFileLoader(id);
        L.localFontData = fd;
        L.owners = [];
        L.LoadFontAsync = function (basePath, callback) {
            if (callback) this.callback = callback;
            if (-1 !== this.Status) return true;
            this.Status = 2;
            var me = this;
            fd.blob().then(function (b) { return b.arrayBuffer(); }).then(function (buf) {
                var data = new Uint8Array(buf);
                var face = faceIndexOf(data, fd);
                me.owners.forEach(function (o) { o.info['faceIndex' + o.slot] = face; });
                me.LoadFontFromData(data);   // Status = 0
                if (me.callback) me.callback();
                if (me['externalCallback']) me['externalCallback']();
            }).catch(function (e) {
                if (window.console) console.warn('[localfonts] cannot read ' + fd.postscriptName, e);
                me.Status = 1;
            });
            return false;
        };
        return L;
    }

    // Register one family (name) with the faces chosen for its slots; returns the new CFontInfo
    function registerFamily(A, name, slots) {
        var idx = { R: -1, I: -1, B: -1, BI: -1 }, byFace = {}, loaders = [];
        SLOTS.forEach(function (s) {
            var fd = slots[s];
            if (!fd) return;
            var key = fd.postscriptName + '|' + fd.fullName + '|' + fd.style;
            var L = byFace[key];
            if (!L) {
                L = byFace[key] = makeLoader(A, 'local-font:' + (LF.count + loaders.length) + ':' + fd.postscriptName, fd);
                A.g_font_files.push(L);
                L.index = A.g_font_files.length - 1;
                loaders.push(L);
            }
            idx[s] = L.index;
        });
        var info = new A.CFontInfo(name, 0, idx.R, 0, idx.I, 0, idx.B, 0, idx.BI, 0);
        info.isLocalFont = true;
        loaders.forEach(function (L) {
            SLOTS.forEach(function (s) { if (idx[s] === L.index) L.owners.push({ info: info, slot: s }); });
        });
        A.g_font_infos.push(info);
        A.g_map_font_index[name] = A.g_font_infos.length - 1;
        local.byName[name] = { name: name };
        local.byNorm[norm(name)] = local.byNorm[norm(name)] || { name: name };
        LF.count++;
        return info;
    }

    // Register the fonts of this computer that the server does not have; returns the names added
    LF.register = function () {
        var A = engine(), list = LF.fonts;
        if (!A || !list) return [];
        var fams = {}, order = [];
        list.forEach(function (fd) {
            var f = fams[fd.family];
            if (!f) { f = fams[fd.family] = []; order.push(fd.family); }
            f.push(fd);
        });
        var added = [];
        var add = function (name, slots) {
            if (!name || local.byName[name] || A.g_map_font_index[name] !== undefined || !slots.R) return;
            registerFamily(A, name, slots);
            added.push(name);
        };
        order.forEach(function (fam) { add(fam, pickSlots(fams[fam])); });
        // weight/width variants addressed by their full name ("Calibri Light", "Segoe UI Semibold")
        list.forEach(function (fd) {
            var full = fd.fullName;
            if (!full || full === fd.family || fams[full] || local.byName[full]) return;
            var c = classify(fd.style);
            if (c.plain && !c.alt) return;                    // plain regular/bold/italic: part of its family
            var base = c.other || c.alt;                      // a variant face
            if (!base || c.italic || c.bold) return;
            var byFull = {};
            list.forEach(function (o) { if (o.family === fd.family) byFull[o.fullName.toLowerCase()] = o; });
            add(full, {
                R: fd,
                I: byFull[(full + ' Italic').toLowerCase()] || null,
                B: byFull[(full + ' Bold').toLowerCase()] || null,
                BI: byFull[(full + ' Bold Italic').toLowerCase()] || null
            });
        });
        return added;
    };

    // Read the list of the fonts of this computer (metadata only, no font file). Must be called from a user
    // gesture the first time (the browser asks its permission then).
    LF.readList = function () {
        if (LF.fonts) return Promise.resolve(LF.fonts);
        if (LF._listing) return LF._listing;
        var bridge = LF.desktopBridge();
        var list = bridge ? bridge.fontConsent().then(function (allowed) {
            if (!allowed) { var error = new Error('Computer font access cancelled'); error.name = 'NotAllowedError'; throw error; }
            return bridge.fonts();
        }).then(function (faces) {
            return faces.map(function (face) {
                return Object.assign({}, face, { blob: function () {
                    return bridge.font(face.id).then(function (bytes) { return new window.Blob([bytes]); });
                } });
            });
        }) : window.queryLocalFonts();
        LF._listing = list.then(function (l) {
            LF._listing = null;
            // a browser that is not allowed to answer may give an empty list instead of an error
            if (!l || !l.length) { var e = new Error('no fonts'); e.name = 'NotAllowedError'; throw e; }
            LF.fonts = l;
            return l;
        }, function (e) { LF._listing = null; throw e; });
        return LF._listing;
    };

    // ---- relayout after a late registration ------------------------------------------------------------
    // Word: every run measures again with the font now behind its name, then the pages are laid out again.
    LF.relayout = function (api) {
        var A = engine();
        try {
            A && A.g_fontApplication && (A.g_fontApplication.FontPickerMap = {});
            var ld = api && api.WordControl && api.WordControl.m_oLogicDocument;
            if (ld && ld.GetAllParagraphs && ld.RecalculateFromStart) {
                ld.GetAllParagraphs({ All: true }).forEach(function (p) {
                    p.Recalc_CompiledPr();
                    p.Recalc_RunsCompiledPr();
                });
                if (AscCommon.g_fontManager2 && AscCommon.g_fontManager2.ClearFontsRasterCache) AscCommon.g_fontManager2.ClearFontsRasterCache();
                api.WordControl.m_oDrawingDocument.ClearCachePages();
                ld.RecalculateFromStart(true, true);
                return true;
            }
        } catch (e) { if (window.console) console.warn('[localfonts] relayout', e); }
        return false;
    };

    // ---- gui font list ---------------------------------------------------------------------------------
    // The names (and ids) the font box needs for the fonts just registered
    LF.guiFonts = function (names) {
        var A = engine();
        return names.map(function (n) { return new A.CFont(n, '', 0); });
    };

    // ---- install the engine hooks (idempotent) -----------------------------------------------------------
    // The editor page loads the engine and opens the document within a few milliseconds, so the loader is
    // wrapped at the very moment it is created (a setter on AscCommon.g_font_loader), not by polling.
    var appWrapped = false;
    function installApp() {
        if (appWrapped) return true;
        var A = engine();
        if (!A) return false;
        appWrapped = true;
        var FA = A.g_fontApplication;
        ['GetFontFileWeb', 'GetFontFile'].forEach(function (fn) {
            var orig = FA[fn];
            if (typeof orig !== 'function') return;
            FA[fn] = function (name) {
                var n = local.byName[name] ? name : LF.resolveName(name);
                if (n && A.g_map_font_index[n] !== undefined && A.g_font_infos[A.g_map_font_index[n]].isLocalFont) return { m_wsFontName: n };
                return orig.apply(this, arguments);
            };
        });
        return true;
    }

    function wrapLoader(loader) {
        if (loader.__elixirLocalFonts) return;
        loader.__elixirLocalFonts = true;
        trace('loader wrapped');
        var origLoad = loader.LoadDocumentFonts;
        loader.LoadDocumentFonts = function (fonts) {
            var args = arguments, me = this, names = [];
            try {
                for (var i in fonts) { var nm = fonts[i] && fonts[i].name; if (nm && names.indexOf(nm) < 0) names.push(nm); }
            } catch (e) { /* no fonts list (pdf) */ }
            LF.docFonts = names;
            installApp();
            trace('LoadDocumentFonts ' + names.length + ' consent=' + LF.hasConsent() + ' support=' + LF.support());
            if (!LF.hasConsent() || LF.support() !== 'ok' || !engine()) { changed(); return origLoad.apply(me, args); }
            // consent given earlier: register the fonts of this computer before the document fonts are loaded
            var started = false, go = function () { if (!started) { started = true; origLoad.apply(me, args); } };
            var timer = setTimeout(go, 8000);
            LF.permission().then(function (p) {
                trace('permission ' + p);
                if (p !== 'granted') { LF.setStatus('off'); return; }
                LF.setStatus('loading');
                return LF.readList().then(function () {
                    trace('registered ' + LF.register().length);
                    LF.setStatus('on');
                });
            }).catch(function (e) {
                if (window.console) console.warn('[localfonts]', e);
                LF.setStatus('error');
            }).then(function () { clearTimeout(timer); go(); });
        };
    }

    LF.install = function () {
        var AC = window.AscCommon;
        if (AC && AC.g_font_loader) wrapLoader(AC.g_font_loader);
        return !!(AC && AC.g_font_loader && installApp());
    };

    // wrap the loader as soon as the engine creates it
    (function trap() {
        var AC = window.AscCommon = window.AscCommon || {};
        if (AC.g_font_loader) return LF.install();
        try {
            Object.defineProperty(AC, 'g_font_loader', {
                configurable: true, enumerable: true,
                get: function () { return undefined; },
                set: function (v) {
                    Object.defineProperty(AC, 'g_font_loader', { configurable: true, enumerable: true, writable: true, value: v });
                    wrapLoader(v);
                }
            });
        } catch (e) { /* the polling below still installs it */ }
    })();

    // ---- user action -----------------------------------------------------------------------------------
    // Called from a click: ask the browser, remember the consent, register, load what the document needs.
    // Resolves {added, found, missing}; rejects with {code: 'insecure'|'unsupported'|'denied'|'blocked'|'error'}
    LF.enable = function (api) {
        var sup = LF.support();
        if (sup !== 'ok') return Promise.reject({ code: sup });
        LF.setStatus('loading');
        return LF.readList().then(function () {
            LF.setConsent(true);
            var before = LF.count;
            var added = LF.register();
            LF.setStatus('on');
            var A = engine();
            if (A && api && added.length && api.sync_InitEditorFonts) api.sync_InitEditorFonts(LF.guiFonts(added));
            var found = (LF.docFonts || []).filter(function (n) { return LF.isLocal(n); });
            var missing = LF.neededNames();
            return new Promise(function (resolve) {
                if (!found.length) return resolve({ added: added, found: found, missing: missing, relayout: false });
                AscCommon.g_font_loader.LoadFonts(found, function () {
                    var ok = LF.relayout(api);
                    resolve({ added: added, found: found, missing: missing, relayout: ok });
                });
            });
        }, function (e) {
            return LF.permission().then(function (p) {
                var code = (p === 'denied' || (e && e.name === 'NotAllowedError')) ? 'denied' : 'blocked';
                LF.setStatus(code);
                throw { code: code, error: e };
            });
        });
    };

    LF.disable = function () {
        var bridge = LF.desktopBridge();
        if (bridge) {
            LF.fonts = null;
            if (typeof bridge.fontDisable === 'function') bridge.fontDisable().catch(function () { LF.setStatus('error'); });
        }
        LF.setConsent(false);
        LF.setStatus('off');
    };

    // fallback: the application part (GetFontFileWeb...) once the engine exists
    var tries = 0, poll = setInterval(function () { if (LF.install() || ++tries > 2400) clearInterval(poll); }, 25);

    return LF;
});

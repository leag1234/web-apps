/*
 * SPDX-FileCopyrightText: 2026 ElixirOffice contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 *
 * Elixir AI, Word toolset.
 *
 * Every tool is a small deterministic function over the public document builder API
 * (the `Api` object a plugin sees in callCommand). The functions below are serialized
 * with toString() and executed through api.callCommand(), the editor's internal analog
 * of a plugin's callCommand: same queue, same sandbox (AscCommon.safePluginEval), same
 * history handling. They must therefore not use closures: arguments are embedded as
 * JSON and the shared helpers are passed as the second parameter.
 *
 * A few read-only layout facts (page of a paragraph, position on the page, drawing wrap)
 * have no public getter; they are read from the builder objects' internal fields inside
 * try/catch and documented in docs/ai.md (they never write).
 *
 * Every tool returns {ok:true, done:{...}} on success (the host turns `done` into the plain
 * words shown to the user) or a structured error {ok:false, error:{code, message, ...}}.
 */

define([], function () {
    'use strict';

    var ElixirAI = (window.Common = window.Common || {}).ElixirAI = window.Common.ElixirAI || {};
    ElixirAI.toolsets = ElixirAI.toolsets || {};

    var ASSISTANT_NAME = 'Elixir AI';
    var MAX_BLOCKS = 400;
    var MAX_TEXT = 400;

    // ---------------------------------------------------------------------------------
    // Code executed inside the editor sandbox
    // ---------------------------------------------------------------------------------

    /* eslint-disable no-undef */
    function HELPERS() {
        var doc = Api.GetDocument();
        var H = { doc: doc };
        H.MM_TW = 56.6929;                       // twips per millimetre
        H.MM_EMU = 36000;                        // EMU per millimetre
        H.count = function () { return doc.GetElementsCount(); };
        H.el = function (i) { return doc.GetElement(i); };
        H.kind = function (el) { return el ? el.GetClassType() : null; };
        H.round = function (v, d) { var k = Math.pow(10, d || 0); return Math.round(v * k) / k; };
        H.styleName = function (p) {
            var s = p.GetStyle();
            return s ? s.GetName() : 'Normal';
        };
        // Text as it reads once the tracked changes are accepted (deleted runs left out): the model
        // must not see "old+new" after a tracked replacement
        H.text = function (el) {
            var k = H.kind(el);
            if (k === 'paragraph') {
                var t = '', n = el.GetElementsCount(), any = false;
                for (var i = 0; i < n; i++) {
                    var e = el.GetElement(i);
                    try { if (e.Run && e.Run.GetReviewType && e.Run.GetReviewType() === 1) { any = true; continue; } } catch (x) {}
                    try { t += e.GetText(); } catch (x) {}
                }
                if (!any) t = el.GetText({ Numbering: false, Math: false, NewLineSeparator: ' ', TabSymbol: '\t' });
                return t.replace(/[\r\n]+$/, '');
            }
            if (k === 'table') {
                var rows = [];
                for (var r = 0; r < el.GetRowsCount(); r++) {
                    var row = el.GetRow(r), cells = [];
                    for (var c = 0; c < row.GetCellsCount(); c++) {
                        var cc = row.GetCell(c).GetContent(), ct = [];
                        for (var q = 0; q < cc.GetElementsCount(); q++) { var ce = cc.GetElement(q); if (ce.GetClassType() === 'paragraph') ct.push(H.text(ce)); }
                        cells.push(ct.join(' ').trim());
                    }
                    rows.push(cells);
                }
                return rows;
            }
            if (k === 'blockLvlSdt') {
                try { return el.GetContent().GetText().replace(/[\r\n]+/g, ' ').trim(); } catch (e) { return ''; }
            }
            return '';
        };
        // Effective (style + direct) font size and bold of the first run with text
        H.effective = function (p) {
            var o = { bold: false, size: null };
            try {
                var c = p.Paragraph.Content;
                for (var k = 0; k < c.length; k++) {
                    if (!c[k].Get_CompiledPr || !c[k].GetText || !String(c[k].GetText() || '').trim()) continue;
                    var pr = c[k].Get_CompiledPr(false);
                    o.bold = !!pr.Bold; o.size = pr.FontSize; break;
                }
            } catch (e) {}
            return o;
        };
        H.clip = function (t, max) { return (t.length > max) ? t.slice(0, max) + '…' : t; };
        // List kind of a paragraph: null, "bullet" or "numbered" (internal numbering format, read only)
        H.listKind = function (p) {
            var lvl = p.GetNumbering();
            if (!lvl) return null;
            try {
                var np = p.Paragraph.GetNumPr();
                var num = p.Paragraph.Parent.GetNumbering().GetNum(np.NumId);
                var l = num.GetLvl(np.Lvl || 0);
                return (l.IsBulleted && l.IsBulleted()) ? 'bullet' : 'numbered';
            } catch (e) { return 'numbered'; }
        };
        H.wrapOf = function (d) {
            try {
                var pd = d.getParaDrawing ? d.getParaDrawing() : d.Drawing;
                if (pd.Is_Inline()) return 'inline';
                var w = pd.wrappingType;
                if (w === 1) return 'square';
                if (w === 2) return 'through';
                if (w === 3) return 'tight';
                if (w === 4) return 'topAndBottom';
                return pd.behindDoc ? 'behind' : 'inFront';
            } catch (e) { return null; }
        };
        H.halignOf = function (d) {
            try {
                var pd = d.getParaDrawing ? d.getParaDrawing() : d.Drawing;
                if (pd.Is_Inline()) return null;
                var ph = pd.PositionH;
                if (ph && ph.Align) return ['center', 'inside', 'left', 'outside', 'right'][ph.Value] || String(ph.Value);
                return null;
            } catch (e) { return null; }
        };
        H.drawings = function (p) {
            var out = [];
            try {
                var all = p.GetAllDrawingObjects();
                for (var i = 0; i < all.length; i++) {
                    var d = all[i];
                    out.push({ n: i, kind: d.GetClassType(), w: H.round(d.GetWidth() / H.MM_EMU, 1), h: H.round(d.GetHeight() / H.MM_EMU, 1), wrap: H.wrapOf(d), halign: H.halignOf(d) });
                }
            } catch (e) {}
            return out;
        };
        H.block = function (i, max) {
            var el = H.el(i), k = H.kind(el), b = { i: i, type: k };
            if (k === 'paragraph') {
                b.style = H.styleName(el);
                b.text = H.clip(H.text(el), max);
                var del = false;
                try { del = el.Paragraph.GetReviewType() === 1; } catch (e) {}
                if (del && !b.text) b.deleted = true;   // a tracked deletion, gone once accepted
                else if (!b.text) b.empty = true;
                var lk = H.listKind(el);
                if (lk) {
                    b.list = lk;
                    try { b.level = el.GetNumbering().GetLevelIndex(); } catch (e) {}
                }
                try {
                    var pp = el.GetParaPr(), jc = pp.GetJc();
                    if (jc && jc !== 'left') b.align = jc === 'both' ? 'justify' : jc;
                } catch (e) {}
                var ef = H.effective(el);
                if (ef.bold) b.bold = true;
                if (ef.size) b.size = ef.size;
                var sec = el.GetSection();
                if (sec && sec.GetColumnsCount() > 1) b.columns = sec.GetColumnsCount();
                var ds = H.drawings(el);
                if (ds.length) b.drawings = ds;
            } else if (k === 'table') {
                var rows = H.text(el);
                b.rows = rows.length;
                b.cols = rows.length ? rows[0].length : 0;
                b.cells = rows.slice(0, 6).map(function (r) { return r.map(function (c) { return H.clip(c, 60); }); });
                if (rows.length > 6) b.more_rows = rows.length - 6;
                try { if (el.GetRow(0).Row.IsHeader()) b.header_row = true; } catch (e) {}
            } else if (k === 'blockLvlSdt') {
                b.text = H.clip(H.text(el), 120);
                try { if (el.Sdt.IsBuiltInTableOfContents()) b.toc = true; } catch (e) {}
            }
            return b;
        };
        // Index of the top-level body element that contains paragraph p
        H.topIndex = function (p) {
            var t = p.GetParentTable();
            if (t) {
                while (t.GetParentTable()) t = t.GetParentTable();
                return t.GetPosInParent();
            }
            var cc = p.GetParentContentControl();
            if (cc && cc.GetPosInParent) return cc.GetPosInParent();
            return p.GetPosInParent();
        };
        // Range from the start of block `from` to the end of block `to`
        H.range = function (from, to) {
            var a = H.el(from).GetRange();
            if (to > from) a = a.ExpandTo(H.el(to).GetRange());
            return a;
        };
        // Paragraphs of blocks from..to (the paragraphs of a table block are included)
        H.paras = function (from, to) {
            var out = [];
            for (var i = from; i <= to; i++) {
                var el = H.el(i), k = H.kind(el);
                if (k === 'paragraph') out.push({ block: i, p: el });
                else if (k === 'table' || k === 'blockLvlSdt') {
                    var ps = el.GetRange().GetAllParagraphs();
                    for (var j = 0; j < ps.length; j++) out.push({ block: i, p: ps[j] });
                }
            }
            return out;
        };
        H.paragraphStyles = function () {
            return doc.GetAllStyles().filter(function (s) { return s.GetType() === 'paragraph'; })
                .map(function (s) { return s.GetName(); })
                .filter(function (n) { return !/^(toc \d|footnote text|endnote text|table of figures|Header|Footer|Caption|TOC Heading)$/i.test(n); });
        };
        H.tableStyles = function () {
            return doc.GetAllStyles().filter(function (s) { return s.GetType() === 'table'; }).map(function (s) { return s.GetName(); });
        };
        H.revisions = function () {
            var rep = doc.GetReviewReport(), n = 0, byAuthor = {};
            for (var a in rep) { if (!Object.prototype.hasOwnProperty.call(rep, a) || !Array.isArray(rep[a])) continue; byAuthor[a] = rep[a].length; n += rep[a].length; }
            return { count: n, byAuthor: byAuthor };
        };
        H.fail = function (code, message, extra) {
            var e = { ok: false, error: { code: code, message: message } };
            for (var k in extra || {}) e.error[k] = extra[k];
            return e;
        };
        H.paraStyle = function (name) {
            var st = doc.GetStyle(name);
            return (st && st.GetType() === 'paragraph' && st.GetName() === name) ? st : null;
        };
        // Apply a paragraph style to the paragraphs of blocks from..to (tables are skipped)
        H.applyStyle = function (from, to, style, clear) {
            var st = H.paraStyle(style);
            if (!st) return H.fail('unknown_style', 'No paragraph style named "' + style + '".', { available: H.paragraphStyles() });
            var done = 0;
            for (var i = from; i <= to; i++) {
                var el = H.el(i);
                if (H.kind(el) !== 'paragraph') continue;
                // ApiRange.SetStyle is recorded as a paragraph-properties revision; ApiParagraph.SetStyle is not.
                el.GetRange().SetStyle(st);
                if (clear) H.clearDirect(el.GetRange());
                done++;
            }
            return { ok: true, styled_paragraphs: done, style: style };
        };
        H.clearDirect = function (range) {
            var tp = Api.CreateTextPr();
            tp.SetBold(null); tp.SetItalic(null); tp.SetUnderline(null); tp.SetStrikeout(null); tp.SetCaps(null); tp.SetSmallCaps(null);
            range.SetTextPr(tp);
        };
        // Occurrences of `find` in paragraph p, optionally whole words only (ApiRange[] in document order)
        H.search = function (p, find, matchCase, wholeWord) {
            var ranges = p.Search(find, !!matchCase) || [];
            if (!wholeWord || !ranges.length) return ranges;
            var text = H.text(p), hay = matchCase ? text : text.toLowerCase(), needle = matchCase ? find : find.toLowerCase();
            var keep = [], pos = 0, k = 0, isWord = function (ch) { return !!ch && /[0-9A-Za-zÀ-ɏͰ-ϿЀ-ӿ_]/.test(ch); };
            while ((pos = hay.indexOf(needle, pos)) >= 0) {
                if (!isWord(text.charAt(pos - 1)) && !isWord(text.charAt(pos + needle.length))) keep.push(k);
                k++; pos += Math.max(1, needle.length);
            }
            if (k !== ranges.length) return ranges; // could not map the occurrences: keep them all
            return keep.map(function (i) { return ranges[i]; });
        };
        H.sections = function () {
            var out = [], secs = doc.GetSections();
            for (var i = 0; i < secs.length; i++) {
                var s = secs[i], w = s.GetPageWidth(), h = s.GetPageHeight();
                var mm = function (tw) { return H.round(tw / H.MM_TW, 1); };
                var paper = null, W = Math.min(w, h) / H.MM_TW, Hh = Math.max(w, h) / H.MM_TW;
                var sizes = { A3: [297, 420], A4: [210, 297], A5: [148, 210], A6: [105, 148], Letter: [215.9, 279.4], Legal: [215.9, 355.6] };
                for (var nm in sizes) if (Math.abs(sizes[nm][0] - W) < 2 && Math.abs(sizes[nm][1] - Hh) < 2) paper = nm;
                out.push({
                    section: i, paper: paper, orientation: w > h ? 'landscape' : 'portrait', width_mm: mm(w), height_mm: mm(h),
                    margins_mm: { top: mm(s.GetPageMarginTop()), bottom: mm(s.GetPageMarginBottom()), left: mm(s.GetPageMarginLeft()), right: mm(s.GetPageMarginRight()) },
                    text_width_mm: mm(w - s.GetPageMarginLeft() - s.GetPageMarginRight()),
                    columns: s.GetColumnsCount(), type: s.GetType()
                });
            }
            return out;
        };
        // Header / footer text, with the page number fields shown as {PAGE} and {NUMPAGES}
        H.hdrFtrText = function (sec, part) {
            try {
                var c = part === 'header' ? sec.GetHeader('default', false) : sec.GetFooter('default', false);
                if (!c) return '';
                var t = c.GetText().replace(/[\r\n]+/g, ' ').trim(), j = '';
                try { j = JSON.stringify(c.ToJSON()); } catch (e) {}
                if (/pgNum|"instr":"\s*PAGE\b/.test(j)) t += ' {PAGE}';
                if (/pagesCnt|NUMPAGES/.test(j)) t += ' {NUMPAGES}';
                return t.trim();
            } catch (e) { return ''; }
        };
        // Elements wider than their text column (drawings, tables), measured in mm
        H.overflow = function () {
            var out = [], secs = H.sections(), n = H.count();
            for (var i = 0; i < n; i++) {
                var el = H.el(i), k = H.kind(el), sec;
                try { sec = el.GetSection ? el.GetSection() : null; } catch (e) { sec = null; }
                var colW = secs[secs.length - 1].text_width_mm;
                try {
                    var all = doc.GetSections();
                    for (var s = 0; s < all.length; s++) if (sec && all[s].Section === sec.Section) colW = secs[s].text_width_mm / Math.max(1, secs[s].columns);
                } catch (e) {}
                if (k === 'paragraph') {
                    var ds = H.drawings(el);
                    for (var j = 0; j < ds.length; j++) if (ds[j].w > colW + 0.5) out.push({ block: i, drawing: j, width_mm: ds[j].w, column_mm: H.round(colW, 1) });
                } else if (k === 'table') {
                    try {
                        // laid-out width (the layout of the last command); the grid if there is no layout yet
                        var tw = 0, b = el.Table.GetPageBounds ? el.Table.GetPageBounds(0) : null;
                        if (b) tw = b.Right - b.Left;
                        if (!tw) { var g = el.Table.TableGrid || []; for (var gi = 0; gi < g.length; gi++) tw += g[gi]; }
                        if (tw > colW + 1) out.push({ block: i, table: true, width_mm: H.round(tw, 1), column_mm: H.round(colW, 1) });
                    } catch (e) {}
                }
            }
            return out;
        };
        // Shrinks the drawings and tables wider than their text column
        H.fit = function () {
            var list = H.overflow(), n = 0;
            for (var i = 0; i < list.length; i++) {
                var o = list[i], el = H.el(o.block);
                if (o.table) {
                    el.SetWidth('twips', Math.round((o.column_mm - 1) * H.MM_TW));
                    for (var r = 0; r < el.GetRowsCount(); r++) {
                        var row = el.GetRow(r), nc = row.GetCellsCount();
                        for (var c = 0; c < nc; c++) row.GetCell(c).SetWidth('twips', Math.floor((o.column_mm - 1) * H.MM_TW / nc));
                    }
                    n++; continue;
                }
                var d = el.GetAllDrawingObjects()[o.drawing];
                if (!d) continue;
                var k = (o.column_mm - 1) / o.width_mm;
                d.SetSize(Math.round(d.GetWidth() * k), Math.round(d.GetHeight() * k));
                n++;
            }
            return { n: n, list: list };
        };
        return H;
    }

    var CMD = {};

    // Initial context: selection -> block range, or the whole document
    CMD.init = function (a, H) {
        var total = H.count(), from = 0, to = total - 1, scope = 'document', selText = '', image = null, cursor = null;
        var r = H.doc.GetRangeBySelect();
        if (r && r.GetText() && r.GetText().replace(/[\s\r\n]/g, '').length > 0) {
            var paras = r.GetAllParagraphs();
            if (paras.length) {
                from = H.topIndex(paras[0]);
                to = H.topIndex(paras[paras.length - 1]);
                scope = 'selection';
                selText = r.GetText();
            }
        }
        try {
            var sd = H.doc.GetSelectedDrawings();
            if (sd && sd.length) {
                var pp = sd[0].GetParentParagraph();
                if (pp) {
                    var bi = H.topIndex(pp), all = pp.GetAllDrawingObjects(), n = 0;
                    for (var i = 0; i < all.length; i++) if (all[i].Drawing === sd[0].Drawing) n = i;
                    image = { block: bi, drawing: n };
                    if (scope === 'document') { from = to = bi; scope = 'selection'; }
                }
            }
        } catch (e) {}
        try {
            var cp = H.doc.GetCurrentParagraph();
            if (cp) cursor = H.topIndex(cp);
        } catch (e) {}
        return {
            ok: true, scope: scope, from: from, to: to, total: total, styles: H.paragraphStyles(), tableStyles: H.tableStyles(),
            revisions: H.revisions(), selectedText: selText.length <= 3000 ? selText : selText.slice(0, 3000), image: image, cursor: cursor,
            sections: H.sections()
        };
    };

    CMD.read_context = function (a, H) {
        var blocks = [], last = Math.min(a.to, H.count() - 1), truncated = false;
        for (var i = a.from; i <= last; i++) {
            if (blocks.length >= a.maxBlocks) { truncated = true; break; }
            blocks.push(H.block(i, a.maxText));
        }
        var out = { ok: true, scope: a.scope, from: a.from, to: last, blocks: blocks };
        if (truncated) out.truncated = true;
        return out;
    };

    CMD.begin = function (a, H) {
        H.doc.SetAssistantTrackRevisions(true, a.name);
        return { ok: true, track: H.doc.IsTrackRevisions() };
    };

    CMD.end = function (a, H) {
        H.doc.SetAssistantTrackRevisions(false);
        return { ok: true, revisions: H.revisions() };
    };

    CMD.revisions = function (a, H) {
        return { ok: true, revisions: H.revisions() };
    };

    CMD.select = function (a, H) {
        var last = Math.min(a.to, H.count() - 1);
        H.range(a.from, last).Select();
        return { ok: true };
    };

    // ---- reading ----

    CMD.document_info = function (a, H) {
        var st = H.doc.GetStatistics(), out = { ok: true };
        out.pages = H.doc.GetPageCount(); out.words = st.WordsCount; out.paragraphs = st.ParagraphCount; out.characters = st.SymbolsWSCount; out.characters_no_spaces = st.SymbolsCount;
        if (a.selectedText) {
            var w = a.selectedText.split(/\s+/).filter(function (x) { return /[0-9A-Za-zÀ-ɏЀ-ӿ぀-鿿]/.test(x); });
            out.selection = { words: w.length, characters: a.selectedText.replace(/[\r\n]/g, '').length };
        }
        var outline = [], n = H.count();
        for (var i = 0; i < n && outline.length < 80; i++) {
            var el = H.el(i);
            if (H.kind(el) !== 'paragraph') continue;
            var s = H.styleName(el), m = /^Heading (\d)$/.exec(s);
            if (m || s === 'Title') outline.push({ block: i, level: m ? +m[1] : 0, text: H.clip(H.text(el), 80) });
        }
        out.outline = outline;
        out.sections = H.sections();
        var secs = H.doc.GetSections();
        out.header = H.hdrFtrText(secs[0], 'header');
        out.footer = H.hdrFtrText(secs[0], 'footer');
        out.overflowing = H.overflow();
        out.comments = (H.doc.GetAllComments() || []).length;
        return out;
    };

    CMD.find_text = function (a, H) {
        var hits = [], total = 0, ps = H.paras(a.from, a.to);
        for (var i = 0; i < ps.length; i++) {
            var n = H.search(ps[i].p, a.text, a.match_case, a.whole_word).length;
            if (!n) continue;
            total += n;
            hits.push({ block: ps[i].block, count: n, text: H.clip(H.text(ps[i].p), 120) });
        }
        return { ok: true, total: total, blocks: hits.slice(0, 60), done: { n: total, find: a.text } };
    };

    CMD.explain_layout = function (a, H) {
        var el = H.el(a.block), out = { ok: true, block: a.block, type: H.kind(el) };
        var page = function (x) { try { return x.GetRange().GetStartPage() + 1; } catch (e) { return null; } };
        out.page = page(el);
        if (a.block > 0) out.previous_block_page = page(H.el(a.block - 1));
        var secs = H.sections(), s0 = secs[0];
        try {
            var all = H.doc.GetSections(), sec = el.GetSection ? el.GetSection() : null;
            for (var s = 0; s < all.length; s++) if (sec && all[s].Section === sec.Section) s0 = secs[s];
        } catch (e) {}
        out.page_setup = s0;
        out.text_height_mm = H.round(s0.height_mm - s0.margins_mm.top - s0.margins_mm.bottom, 1);
        if (H.kind(el) === 'paragraph') {
            try {
                var j = JSON.parse(el.ToJSON(false, false)), pPr = j.pPr || {};
                out.keep_with_next = !!pPr.keepNext; out.keep_lines = !!pPr.keepLines; out.page_break_before = !!pPr.pageBreakBefore;
            } catch (e) {}
            out.drawings = H.drawings(el);
            try {
                var b = el.Paragraph.GetPageBounds(0);
                out.top_on_page_mm = H.round(b.Top, 1);
                out.paragraph_height_on_first_page_mm = H.round(b.Bottom - b.Top, 1);
            } catch (e) {}
            if (a.block > 0) {
                try {
                    var prev = H.el(a.block - 1), pb = null;
                    if (H.kind(prev) === 'paragraph') pb = prev.Paragraph.GetPageBounds(prev.Paragraph.GetPagesCount() - 1);
                    else if (H.kind(prev) === 'table') pb = prev.Table.GetPageBounds(prev.Table.GetPagesCount() - 1);
                    if (pb) {
                        out.previous_block_bottom_mm = H.round(pb.Bottom, 1);
                        out.space_left_after_previous_block_mm = H.round(s0.height_mm - s0.margins_mm.bottom - pb.Bottom, 1);
                    }
                } catch (e) {}
            }
        }
        out.done = { block: a.block };
        return out;
    };

    // Lines that look like titles, with their features and a suggested outline level
    CMD.get_outline_candidates = function (a, H) {
        var info = [], sizes = {}, n = Math.min(a.to, H.count() - 1);
        var eff = function (p) {
            var o = { bold: false, size: null };
            try {
                var c = p.Paragraph.Content;
                for (var k = 0; k < c.length; k++) {
                    if (!c[k].Get_CompiledPr || !c[k].GetText || !String(c[k].GetText() || '').trim()) continue;
                    var pr = c[k].Get_CompiledPr(false);
                    o.bold = !!pr.Bold; o.size = pr.FontSize; break;
                }
            } catch (e) {}
            if (o.size === null) {
                try { var tp = p.GetRange().GetTextPr(); o.bold = !!tp.GetBold(); o.size = tp.GetFontSize() || null; } catch (e) {}
            }
            return o;
        };
        for (var i = a.from; i <= n; i++) {
            var el = H.el(i);
            if (H.kind(el) !== 'paragraph') continue;
            var t = H.text(el).trim();
            if (!t) continue;
            var f = eff(el), x = { block: i, text: H.clip(t, 100), style: H.styleName(el), bold: f.bold, size: f.size, length: t.length };
            var m = /^\s*((\d+)((\.\d+)*))[.)]?\s+\S/.exec(t);
            if (m) x.numbering = { pattern: m[1], depth: m[1].split('.').filter(Boolean).length };
            else if (/^\s*[IVXLC]+[.)]\s+\S/.test(t)) x.numbering = { pattern: 'roman', depth: 1 };
            else if (/^\s*[A-Za-z][.)]\s+\S/.test(t)) x.numbering = { pattern: 'letter', depth: 2 };
            var lk = H.listKind(el);
            if (lk) { x.list = lk; try { x.list_level = el.GetNumbering().GetLevelIndex(); } catch (e) {} }
            if (/[A-Za-zÀ-ɏ]/.test(t) && t === t.toUpperCase()) x.caps = true;
            x.ends_with_period = /[.;:,!?]$/.test(t) && !/\?$/.test(t);
            info.push(x);
            if (f.size) sizes[f.size] = (sizes[f.size] || 0) + t.length;
        }
        var body = null, best = -1;
        for (var s in sizes) if (sizes[s] > best) { best = sizes[s]; body = +s; }
        var cands = info.filter(function (x) {
            if (/^Heading \d$|^Title$/.test(x.style)) return true;
            if (x.length > 120 || (x.ends_with_period && !x.numbering)) return false;
            var big = body && x.size && x.size > body + 0.5;
            return big || x.bold || x.caps || (x.numbering && x.length < 90 && !x.ends_with_period);
        });
        // level from the numbering depth when the titles are numbered, else from the size/bold rank
        var numbered = cands.filter(function (x) { return x.numbering && x.numbering.pattern !== 'letter'; }).length >= Math.max(2, cands.length / 2);
        var keys = [];
        cands.forEach(function (x) { var k = (x.size || 0) * 10 + (x.bold ? 1 : 0); if (keys.indexOf(k) < 0) keys.push(k); });
        keys.sort(function (p, q) { return q - p; });
        cands.forEach(function (x, idx) {
            var hm = /^Heading (\d)$/.exec(x.style);
            if (hm) x.suggested_level = +hm[1];
            else if (numbered && x.numbering) x.suggested_level = Math.min(3, x.numbering.depth);
            else x.suggested_level = Math.min(3, keys.indexOf((x.size || 0) * 10 + (x.bold ? 1 : 0)) + 1);
        });
        // a first line clearly bigger than every other candidate is the document title
        if (cands.length > 1 && cands[0].block === info[0].block && cands[0].size && cands.slice(1).every(function (x) { return (x.size || 0) < cands[0].size; }) && !cands[0].numbering) {
            cands[0].suggested_level = 0;
            cands[0].note = 'document title (Title style, not in the outline)';
            for (var j = 1; j < cands.length; j++) if (!numbered && !/^Heading/.test(cands[j].style)) cands[j].suggested_level = Math.max(1, cands[j].suggested_level - 1);
        }
        return { ok: true, body_size: body, numbered_titles: numbered, candidates: cands.slice(0, 120), hint: 'suggested_level: 0 = document title, 1 = Heading 1, 2 = Heading 2… Check it against the text before applying.', done: { n: cands.length } };
    };

    CMD.get_layout_summary = function (a, H) {
        var st = H.doc.GetStatistics(), n = H.count(), tables = 0, drawings = [], big = [];
        for (var i = 0; i < n; i++) {
            var el = H.el(i), k = H.kind(el);
            if (k === 'table') tables++;
            if (k !== 'paragraph') continue;
            var ds = H.drawings(el);
            for (var j = 0; j < ds.length; j++) drawings.push({ block: i, drawing: j, width_mm: ds[j].w, height_mm: ds[j].h, wrap: ds[j].wrap });
            try { var sz = el.GetRange().GetTextPr().GetFontSize(); if (sz && sz >= 20) big.push({ block: i, size: sz }); } catch (e) {}
        }
        return { ok: true, pages: H.doc.GetPageCount(), sections: H.sections(), tables: tables, drawings: drawings.slice(0, 40), large_text: big.slice(0, 20), overflowing: H.overflow(), done: {} };
    };

    CMD.find_overflowing_elements = function (a, H) {
        var o = H.overflow();
        return { ok: true, count: o.length, elements: o, done: { n: o.length } };
    };

    // ---- text ----

    CMD.insert_paragraphs = function (a, H) {
        var pos = a.after + 1;
        for (var k = 0; k < a.items.length; k++) {
            var it = a.items[k];
            var p = Api.CreateParagraph();
            p.AddText(it.text);
            var st = H.doc.GetStyle(it.style || 'Normal');
            if (st && st.GetType() === 'paragraph') p.SetStyle(st);
            if (it.align) p.SetJc(it.align === 'justify' ? 'both' : it.align);
            H.doc.AddElement(pos + k, p);
        }
        return { ok: true, inserted: a.items.length, first_index: pos, last_index: pos + a.items.length - 1, tracked: true, done: { n: a.items.length } };
    };

    CMD.replace_text = function (a, H) {
        var el = H.el(a.block);
        if (H.kind(el) !== 'paragraph') return H.fail('not_a_paragraph', 'Block ' + a.block + ' is not a paragraph.');
        var ranges = [];
        if (a.find) {
            ranges = el.Search(a.find, true) || [];
            if (!ranges.length) return H.fail('not_found', 'Text "' + a.find + '" not found in block ' + a.block + '.', { block_text: H.clip(H.text(el), 400) });
            if (!a.all) ranges = [ranges[0]];
        } else {
            // the text only, not the paragraph mark (deleting the mark would merge the paragraph)
            var full = H.text(el), r0 = full.length ? el.GetRange(0, full.length) : null;
            if (r0 && r0.GetText().replace(/[\r\n]+$/, '') !== full) r0 = (el.Search(full, true) || [])[0] || null;
            if (!full.length) { el.AddText(a.text); return { ok: true, replaced: 1, tracked: true, done: { n: 1 } }; }
            if (!r0) return H.fail('cannot_select', 'The text of block ' + a.block + ' could not be selected; use `find` with a part of it.');
            ranges = [r0];
        }
        // Deletion first, then insertion after it: both are tracked revisions
        for (var i = ranges.length - 1; i >= 0; i--) {
            ranges[i].Delete();
            ranges[i].AddText(a.text, 'after');
        }
        return { ok: true, replaced: ranges.length, tracked: true, done: { n: ranges.length } };
    };

    CMD.find_replace = function (a, H) {
        var ps = H.paras(a.from, a.to), total = 0, blocks = [];
        for (var i = 0; i < ps.length; i++) {
            var ranges = H.search(ps[i].p, a.find, a.match_case, a.whole_word);
            for (var j = ranges.length - 1; j >= 0; j--) {
                ranges[j].Delete();
                if (a.replace) ranges[j].AddText(a.replace, 'after');
            }
            if (ranges.length) { total += ranges.length; if (blocks.indexOf(ps[i].block) < 0) blocks.push(ps[i].block); }
        }
        if (!total) return H.fail('not_found', 'No occurrence of "' + a.find + '" in blocks ' + a.from + '-' + a.to + (a.whole_word ? ' (whole words)' : '') + '.');
        return { ok: true, replaced: total, blocks: blocks, tracked: true, done: { n: total, find: a.find, replace: a.replace } };
    };

    CMD.rewrite_blocks = function (a, H) {
        var styles = [];
        for (var i = a.from; i <= a.to; i++) {
            var el = H.el(i);
            if (H.kind(el) !== 'paragraph') return H.fail('not_a_paragraph', 'Block ' + i + ' is a ' + H.kind(el) + '; rewrite_blocks only replaces paragraphs.');
            if (H.drawings(el).length) return H.fail('has_drawing', 'Block ' + i + ' holds a picture: rewriting it would delete the picture. Check the block indices (read_context).');
            styles.push(H.styleName(el));
        }
        var pos = a.to + 1;
        for (var k = 0; k < a.items.length; k++) {
            var it = a.items[k], p = Api.CreateParagraph();
            p.AddText(it.text);
            var st = H.paraStyle(it.style || styles[Math.min(k, styles.length - 1)]);
            if (st) p.SetStyle(st);
            H.doc.AddElement(pos + k, p);
        }
        var del = H.range(a.from, a.to);
        del = del.ExpandTo(H.el(pos).GetRange(0, 0));
        del.Delete();
        return {
            ok: true, inserted: [pos, pos + a.items.length - 1], deleted_blocks: [a.from, a.to], tracked: true,
            note: 'The new paragraphs are blocks ' + pos + '-' + (pos + a.items.length - 1) + '; the old ones stay visible as tracked deletions (blocks ' + a.from + '-' + a.to + ') until accepted.',
            done: { n: a.items.length }
        };
    };

    CMD.delete_blocks = function (a, H) {
        var del = H.range(a.from, a.to), n = H.count();
        if (a.to + 1 < n) del = del.ExpandTo(H.el(a.to + 1).GetRange(0, 0));
        del.Delete();
        return { ok: true, deleted_blocks: [a.from, a.to], tracked: true, note: 'Deleted blocks stay visible as tracked deletions until accepted; indices do not change.', done: { n: a.to - a.from + 1 } };
    };

    // ---- styles and structure ----

    CMD.apply_style = function (a, H) {
        if (a.blocks) {
            var n = 0;
            for (var k = 0; k < a.blocks.length; k++) {
                var rr = H.applyStyle(a.blocks[k], a.blocks[k], a.style, a.clear_direct_formatting);
                if (!rr.ok) return rr;
                n += rr.styled_paragraphs;
            }
            return { ok: true, styled_paragraphs: n, style: a.style, tracked: true, done: { n: n, style: a.style } };
        }
        var r = H.applyStyle(a.from, a.to, a.style, a.clear_direct_formatting);
        if (r.ok) { r.tracked = true; r.done = { n: r.styled_paragraphs, style: a.style }; }
        return r;
    };

    CMD.apply_styles = function (a, H) {
        var results = [], total = 0, byStyle = {};
        for (var k = 0; k < a.changes.length; k++) {
            var c = a.changes[k];
            var r = H.applyStyle(c.from, c.to, c.style, c.clear_direct_formatting);
            if (!r.ok) { r.error.change = k; r.applied_before_error = results; return r; }
            total += r.styled_paragraphs;
            byStyle[c.style] = (byStyle[c.style] || 0) + r.styled_paragraphs;
            results.push({ from: c.from, to: c.to, style: c.style });
        }
        return { ok: true, styled_paragraphs: total, changes: results.length, tracked: true, done: { n: total, byStyle: byStyle } };
    };

    // Style definition (all paragraphs of the style follow): the Word way to apply a style guide
    CMD.modify_style = function (a, H) {
        var st = H.doc.GetStyle(a.style);
        if (!st || st.GetName() !== a.style) return H.fail('unknown_style', 'No style named "' + a.style + '".', { available: H.paragraphStyles() });
        var tp = st.GetTextPr();
        if (a.font) tp.SetFontFamily(a.font);
        if (a.size_pt) tp.SetFontSize(Math.round(a.size_pt * 2));
        if (a.bold !== undefined) tp.SetBold(a.bold);
        if (a.italic !== undefined) tp.SetItalic(a.italic);
        if (a.color) { var m = /^#?([0-9a-f]{2})([0-9a-f]{2})([0-9a-f]{2})$/i.exec(a.color); if (m) tp.SetColor(parseInt(m[1], 16), parseInt(m[2], 16), parseInt(m[3], 16), false); }
        if (st.GetType() === 'paragraph' && (a.align || a.space_after_pt !== undefined || a.space_before_pt !== undefined)) {
            var pp = st.GetParaPr();
            if (a.align) pp.SetJc(a.align === 'justify' ? 'both' : a.align);
            if (a.space_after_pt !== undefined) pp.SetSpacingAfter(Math.round(a.space_after_pt * 20));
            if (a.space_before_pt !== undefined) pp.SetSpacingBefore(Math.round(a.space_before_pt * 20));
        }
        return { ok: true, style: a.style, tracked: false, note: 'Direct formatting on runs still overrides the style; use format_text to change those runs.', done: { style: a.style } };
    };

    CMD.change_heading_level = function (a, H) {
        var n = 0, missing = [];
        for (var i = a.from; i <= a.to; i++) {
            var el = H.el(i);
            if (H.kind(el) !== 'paragraph') continue;
            var m = /^Heading (\d)$/.exec(H.styleName(el));
            if (!m) continue;
            var lv = Math.max(1, Math.min(9, +m[1] + a.delta)), st = H.paraStyle('Heading ' + lv);
            if (!st) { missing.push('Heading ' + lv); continue; }
            if (lv !== +m[1]) { el.GetRange().SetStyle(st); n++; }
        }
        if (!n) return H.fail('no_heading', 'No heading paragraph (Heading 1-9) to change in blocks ' + a.from + '-' + a.to + (missing.length ? '; missing styles: ' + missing.join(', ') : '') + '.');
        return { ok: true, changed: n, tracked: true, done: { n: n } };
    };

    CMD.format_text = function (a, H) {
        var ranges = [];
        if (a.find) {
            var ps = H.paras(a.from, a.to);
            for (var i = 0; i < ps.length; i++) {
                var rs = H.search(ps[i].p, a.find, a.match_case, a.whole_word);
                for (var j = 0; j < rs.length; j++) ranges.push(rs[j]);
                if (ranges.length && !a.all) break;
            }
            if (!a.all) ranges = ranges.slice(0, 1);
            if (!ranges.length) return H.fail('not_found', 'Text "' + a.find + '" not found in blocks ' + a.from + '-' + a.to + '.');
        } else {
            ranges = [H.range(a.from, a.to)];
        }
        var hex = function (c) { var m = /^#?([0-9a-f]{2})([0-9a-f]{2})([0-9a-f]{2})$/i.exec(c || ''); return m ? [parseInt(m[1], 16), parseInt(m[2], 16), parseInt(m[3], 16)] : null; };
        for (var k = 0; k < ranges.length; k++) {
            var r = ranges[k];
            if (a.bold !== undefined) r.SetBold(a.bold);
            if (a.italic !== undefined) r.SetItalic(a.italic);
            if (a.underline !== undefined) r.SetUnderline(a.underline);
            if (a.strikethrough !== undefined) r.SetStrikeout(a.strikethrough);
            if (a.caps !== undefined) r.SetCaps(a.caps);
            if (a.font) r.SetFontFamily(a.font);
            if (a.size_pt) r.SetFontSize(Math.round(a.size_pt * 2) / 2); // ApiRange.SetFontSize takes points (its JSDoc says half-points)
            if (a.color) { var c = hex(a.color); if (c) r.SetColor(c[0], c[1], c[2]); }
            if (a.highlight) r.SetHighlight(a.highlight);
        }
        return { ok: true, ranges: ranges.length, tracked: true, done: { n: ranges.length } };
    };

    CMD.format_paragraphs = function (a, H) {
        var ps = H.paras(a.from, a.to), tw = function (mm) { return Math.round(mm * H.MM_TW); };
        if (a.paragraphs_only_style) ps = ps.filter(function (x) { return H.styleName(x.p) === a.paragraphs_only_style; });
        for (var i = 0; i < ps.length; i++) {
            var p = ps[i].p;
            if (a.align) p.SetJc(a.align === 'justify' ? 'both' : a.align);
            if (a.indent_left_mm !== undefined) p.SetIndLeft(tw(a.indent_left_mm));
            if (a.indent_right_mm !== undefined) p.SetIndRight(tw(a.indent_right_mm));
            if (a.first_line_mm !== undefined) p.SetIndFirstLine(tw(a.first_line_mm));
            if (a.space_before_pt !== undefined) p.SetSpacingBefore(Math.round(a.space_before_pt * 20));
            if (a.space_after_pt !== undefined) p.SetSpacingAfter(Math.round(a.space_after_pt * 20));
            if (a.line_spacing !== undefined) p.SetSpacingLine(Math.round(a.line_spacing * 240), 'auto');
            if (a.keep_with_next !== undefined) p.SetKeepNext(a.keep_with_next);
            if (a.keep_lines !== undefined) p.SetKeepLines(a.keep_lines);
            if (a.page_break_before !== undefined) p.SetPageBreakBefore(a.page_break_before);
        }
        if (!ps.length) return H.fail('no_paragraph', 'No paragraph in blocks ' + a.from + '-' + a.to + '.');
        return { ok: true, paragraphs: ps.length, tracked: true, done: { n: ps.length } };
    };

    CMD.set_list = function (a, H) {
        var ps = [], idx = [];
        if (a.blocks) idx = a.blocks.slice().sort(function (x, y) { return x - y; });
        else for (var b = a.from; b <= a.to; b++) idx.push(b);
        for (var ii = 0; ii < idx.length; ii++) {
            var i = idx[ii], el = H.el(i);
            if (H.kind(el) !== 'paragraph') continue;
            if (a.skip_empty !== false && !H.text(el).replace(/\s/g, '').length) continue;
            ps.push(el);
        }
        if (!ps.length) return H.fail('no_paragraph', 'No non-empty paragraph in blocks ' + idx.join(',') + '.');
        if (a.kind === 'none') {
            var normal = H.paraStyle('Normal');
            for (var j = 0; j < ps.length; j++) {
                var np = ps[j].GetParaPr();
                // removing the numbering of a paragraph: an empty numbering reference
                try { ps[j].Paragraph.private_AddPrChange(); } catch (e) {}
                if (ps[j].GetNumbering()) {
                    var pr = ps[j].GetParaPr();
                    pr.ParaPr.NumPr = new AscWord.NumPr(0, 0);
                    pr.private_OnChange();
                }
                if (normal && H.styleName(ps[j]) === 'List Paragraph') ps[j].GetRange().SetStyle(normal);
            }
            return { ok: true, paragraphs: ps.length, tracked: true, done: { n: ps.length, kind: 'none' } };
        }
        var num = H.doc.CreateNumbering(a.kind === 'bullet' ? 'bullet' : 'numbered');
        var lvl = num.GetLevel(a.level || 0);
        if (a.kind === 'numbered' && a.format) lvl.SetTemplateType(a.format);
        if (a.start !== undefined) lvl.SetStart(a.start);
        for (var k = 0; k < ps.length; k++) ps[k].SetNumbering(num.GetLevel(a.level || 0));
        return { ok: true, paragraphs: ps.length, kind: a.kind, tracked: true, done: { n: ps.length, kind: a.kind } };
    };

    CMD.change_list_level = function (a, H) {
        var n = 0;
        for (var i = a.from; i <= a.to; i++) {
            var el = H.el(i);
            if (H.kind(el) !== 'paragraph') continue;
            var lvl = el.GetNumbering();
            if (!lvl) continue;
            var nl = Math.max(0, Math.min(8, lvl.GetLevelIndex() + a.delta));
            if (nl === lvl.GetLevelIndex()) continue;
            el.SetNumbering(lvl.GetNumbering().GetLevel(nl));
            n++;
        }
        if (!n) return H.fail('no_list_item', 'No list item whose level can change in blocks ' + a.from + '-' + a.to + '.');
        return { ok: true, changed: n, tracked: true, done: { n: n } };
    };

    CMD.insert_table_of_contents = function (a, H) {
        var n = H.count(), at = Math.max(0, Math.min(a.before, n - 1)), heads = 0;
        for (var i = 0; i < n; i++) {
            var el = H.el(i);
            if (H.kind(el) === 'paragraph' && /^Heading [1-9]$/.test(H.styleName(el))) heads++;
        }
        if (!heads) return H.fail('no_heading', 'The document has no Heading 1-9 paragraph: apply heading styles to the titles first (apply_styles), then insert the table of contents.');
        var target = H.el(at);
        if (H.kind(target) !== 'paragraph') return H.fail('not_a_paragraph', 'Block ' + at + ' is not a paragraph; choose a paragraph to insert the table of contents before.');
        // the table of contents is inserted at the cursor: the start of the target paragraph
        target.GetRange(0, 0).Select();
        H.doc.AddTableOfContents({ ShowPageNums: a.page_numbers !== false, RightAlgn: true, LeaderType: 'dot', FormatAsLinks: true, BuildFrom: { OutlineLvls: a.levels || 3 }, TocStyle: 'standard' });
        return { ok: true, before: at, headings: heads, levels: a.levels || 3, tracked: true, note: 'The table of contents is a new block at index ' + at + '; later blocks moved by +1 or more: call read_context before more edits.', done: { n: heads } };
    };

    // ---- tables ----

    CMD.list_to_table = function (a, H) {
        var texts = [];
        for (var i = a.from; i <= a.to; i++) {
            var el = H.el(i);
            if (H.kind(el) !== 'paragraph')
                return H.fail('not_a_paragraph', 'Block ' + i + ' is a ' + H.kind(el) + ', expected list paragraphs only.');
            var t = H.text(el).trim();
            if (t.length) texts.push(t);
        }
        if (!texts.length) return H.fail('empty_range', 'The range contains no text.');

        var sep = a.separator;
        if (!sep) {
            var candidates = ['\t', ' | ', '|', ';', ' – ', ' — ', ' - ', ': ', ','];
            for (var c = 0; c < candidates.length && !sep; c++) {
                var n = texts[0].split(candidates[c]).length;
                if (n < 2) continue;
                var same = texts.every(function (t) { return t.split(candidates[c]).length === n; });
                if (same) sep = candidates[c];
            }
            if (!sep) {
                for (c = 0; c < candidates.length && !sep; c++)
                    if (texts.every(function (t) { return t.indexOf(candidates[c]) >= 0; })) sep = candidates[c];
            }
        }
        var rows = texts.map(function (t) {
            return (sep ? t.split(sep) : [t]).map(function (s) { return s.trim(); });
        });
        var hasHeader = false;
        if (Array.isArray(a.header)) { rows.unshift(a.header.map(String)); hasHeader = true; }
        else if (a.header === true || a.header === 'first_item') hasHeader = true;

        var cols = 0;
        rows.forEach(function (r) { cols = Math.max(cols, r.length); });
        // Content past the last column is kept in the last cell
        rows = rows.map(function (r) {
            if (r.length > cols) r = r.slice(0, cols - 1).concat([r.slice(cols - 1).join(sep || ' ')]);
            while (r.length < cols) r.push('');
            return r;
        });

        var table = Api.CreateTable(cols, rows.length);
        table.SetWidth('percent', 100);
        var st = H.doc.GetStyle(a.style || 'Table Grid');
        if (st && st.GetType() === 'table') table.SetStyle(st);
        table.SetTableLook(false, hasHeader, false, false, false, false);
        for (var r = 0; r < rows.length; r++) {
            for (var cc = 0; cc < cols; cc++) {
                var para = table.GetCell(r, cc).GetContent().GetElement(0);
                var run = para.AddText(rows[r][cc]);
                if (hasHeader && r === 0 && run) run.SetBold(true);
            }
        }
        if (hasHeader) table.GetRow(0).SetTableHeader(true);

        // Insert the table before the list, then delete the list as a tracked deletion.
        H.doc.AddElement(a.from, table);
        var first = a.from + 1, last = a.to + 1;
        var del = H.range(first, last);
        if (last + 1 < H.count()) del = del.ExpandTo(H.el(last + 1).GetRange(0, 0));
        del.Delete();
        return {
            ok: true, table_index: a.from, rows: rows.length, cols: cols, header_row: hasHeader,
            separator: sep || null, deleted_blocks: [first, last], tracked: true,
            note: 'The table is block ' + a.from + '; the original list now spans blocks ' + first + '-' + last +
                ' and stays visible as a tracked deletion until accepted. Indices of later blocks moved by +1.',
            done: { rows: rows.length, cols: cols }
        };
    };

    CMD.create_table = function (a, H) {
        var rows = a.rows.map(function (r) { return r.map(String); }), cols = 0;
        rows.forEach(function (r) { cols = Math.max(cols, r.length); });
        rows.forEach(function (r) { while (r.length < cols) r.push(''); });
        var table = Api.CreateTable(cols, rows.length);
        table.SetWidth('percent', 100);
        var st = H.doc.GetStyle(a.style || 'Table Grid');
        if (st && st.GetType() === 'table') table.SetStyle(st);
        table.SetTableLook(false, !!a.header, false, false, false, false);
        for (var r = 0; r < rows.length; r++)
            for (var c = 0; c < cols; c++) {
                var run = table.GetCell(r, c).GetContent().GetElement(0).AddText(rows[r][c]);
                if (a.header && r === 0 && run) run.SetBold(true);
            }
        if (a.header) table.GetRow(0).SetTableHeader(true);
        H.doc.AddElement(a.after + 1, table);
        return { ok: true, table_index: a.after + 1, rows: rows.length, cols: cols, tracked: true, done: { rows: rows.length, cols: cols } };
    };

    // Shared by the table tools: the table at block a.table, or an error
    CMD._table = null;

    CMD.table_rows = function (a, H) {
        var t = H.el(a.table);
        if (H.kind(t) !== 'table') return H.fail('not_a_table', 'Block ' + a.table + ' is not a table.');
        var n = t.GetRowsCount(), count = a.count || 1;
        if (a.row < 0 || a.row >= n) return H.fail('invalid_argument', 'Row ' + a.row + ' does not exist (the table has ' + n + ' rows, 0-based).', { field: 'row' });
        if (a.action === 'insert') {
            var before = a.position === 'before';
            t.AddRows(t.GetCell(a.row, 0), count, before);
            var first = before ? a.row : a.row + 1;
            if (a.cells) {
                for (var r = 0; r < Math.min(count, a.cells.length); r++)
                    for (var c = 0; c < (a.cells[r] || []).length; c++) {
                        var cell = t.GetCell(first + r, c);
                        if (cell) cell.GetContent().GetElement(0).AddText(String(a.cells[r][c]));
                    }
            }
            return { ok: true, inserted_rows: [first, first + count - 1], rows: t.GetRowsCount(), tracked: true, done: { n: count, action: 'insert' } };
        }
        var last = Math.min(n - 1, a.row + count - 1);
        for (var k = last; k >= a.row; k--) t.RemoveRow(t.GetCell(k, 0));
        return { ok: true, deleted_rows: [a.row, last], tracked: true, done: { n: last - a.row + 1, action: 'delete' } };
    };

    CMD.table_add_header_row = function (a, H) {
        var t = H.el(a.table);
        if (H.kind(t) !== 'table') return H.fail('not_a_table', 'Block ' + a.table + ' is not a table.');
        var cols = t.GetRow(0).GetCellsCount();
        if (a.titles.length !== cols) return H.fail('invalid_argument', 'The table has ' + cols + ' columns: give ' + cols + ' titles.', { field: 'titles' });
        t.AddRows(t.GetCell(0, 0), 1, true);
        for (var c = 0; c < cols; c++) {
            var run = t.GetCell(0, c).GetContent().GetElement(0).AddText(String(a.titles[c]));
            if (run) run.SetBold(true);
        }
        t.GetRow(0).SetTableHeader(true);
        var look = null;
        try { look = JSON.parse(t.ToJSON(false, false)).tblPr.tblLook; } catch (e) {}
        look = look || {};
        t.SetTableLook(!!look.firstColumn, true, false, false, !look.noHBand, false);
        return { ok: true, rows: t.GetRowsCount(), tracked: true, done: { n: 1, action: 'insert' } };
    };

    CMD.table_columns = function (a, H) {
        var t = H.el(a.table);
        if (H.kind(t) !== 'table') return H.fail('not_a_table', 'Block ' + a.table + ' is not a table.');
        var cols = t.GetRow(0).GetCellsCount(), count = a.count || 1;
        if (a.column < 0 || a.column >= cols) return H.fail('invalid_argument', 'Column ' + a.column + ' does not exist (the first row has ' + cols + ' cells, 0-based).', { field: 'column' });
        if (a.action === 'insert') {
            var before = a.position === 'before';
            t.AddColumns(t.GetCell(0, a.column), count, before);
            var first = before ? a.column : a.column + 1;
            if (a.cells) {
                for (var r = 0; r < Math.min(t.GetRowsCount(), a.cells.length); r++)
                    for (var c = 0; c < Math.min(count, (a.cells[r] || []).length); c++) {
                        var cell = t.GetCell(r, first + c);
                        if (cell) cell.GetContent().GetElement(0).AddText(String(a.cells[r][c]));
                    }
            }
            return { ok: true, inserted_columns: [first, first + count - 1], tracked: false, done: { n: count, action: 'insert' } };
        }
        var last = Math.min(cols - 1, a.column + count - 1);
        for (var k = last; k >= a.column; k--) t.RemoveColumn(t.GetCell(0, k));
        return { ok: true, deleted_columns: [a.column, last], tracked: false, done: { n: last - a.column + 1, action: 'delete' } };
    };

    CMD.table_format = function (a, H) {
        var t = H.el(a.table);
        if (H.kind(t) !== 'table') return H.fail('not_a_table', 'Block ' + a.table + ' is not a table.');
        if (a.style) {
            var st = H.doc.GetStyle(a.style);
            if (!st || st.GetType() !== 'table') return H.fail('unknown_style', 'No table style named "' + a.style + '".', { available: H.tableStyles() });
            t.SetStyle(st);
        }
        if (a.header_row !== undefined) {
            t.GetRow(0).SetTableHeader(a.header_row);
            var row = t.GetRow(0);
            for (var c = 0; c < row.GetCellsCount(); c++) row.GetCell(c).GetContent().GetRange().SetBold(a.header_row);
        }
        if (a.header_row !== undefined || a.banded_rows !== undefined || a.first_column !== undefined) {
            var look = null;
            try { look = JSON.parse(t.ToJSON(false, false)).tblPr.tblLook; } catch (e) {}
            look = look || {};
            t.SetTableLook(a.first_column !== undefined ? a.first_column : !!look.firstColumn, a.header_row !== undefined ? a.header_row : !!look.firstRow,
                false, false, a.banded_rows !== undefined ? a.banded_rows : !look.noHBand, false);
        }
        if (a.width_percent) t.SetWidth('percent', a.width_percent);
        if (a.align) t.SetJc(a.align);
        if (a.column_widths_mm) {
            t.SetTableLayout('fixed');
            for (var r = 0; r < t.GetRowsCount(); r++) {
                var rw = t.GetRow(r);
                for (var k = 0; k < Math.min(rw.GetCellsCount(), a.column_widths_mm.length); k++)
                    rw.GetCell(k).SetWidth('twips', Math.round(a.column_widths_mm[k] * H.MM_TW));
            }
        }
        return { ok: true, tracked: !a.column_widths_mm, done: {} };
    };

    CMD.table_merge_cells = function (a, H) {
        var t = H.el(a.table);
        if (H.kind(t) !== 'table') return H.fail('not_a_table', 'Block ' + a.table + ' is not a table.');
        var cells = [];
        for (var r = a.from_row; r <= a.to_row; r++)
            for (var c = a.from_column; c <= a.to_column; c++) {
                var cell = t.GetCell(r, c);
                if (!cell) return H.fail('invalid_argument', 'Cell (' + r + ', ' + c + ') does not exist.');
                cells.push(cell);
            }
        if (cells.length < 2) return H.fail('invalid_argument', 'Select at least two cells to merge.');
        var m = t.MergeCells(cells);
        if (!m) return H.fail('merge_failed', 'These cells cannot be merged (they must form a rectangle).');
        return { ok: true, merged: cells.length, tracked: false, done: { n: cells.length } };
    };

    // Rewrites the text of one cell as a tracked deletion + insertion (formatting of the cell kept)
    CMD.table_sort = function (a, H) {
        var t = H.el(a.table);
        if (H.kind(t) !== 'table') return H.fail('not_a_table', 'Block ' + a.table + ' is not a table.');
        var n = t.GetRowsCount(), cols = t.GetRow(0).GetCellsCount();
        if (a.column >= cols) return H.fail('invalid_argument', 'Column ' + a.column + ' does not exist (' + cols + ' columns, 0-based).', { field: 'column' });
        var header = a.has_header;
        if (header === undefined) { header = false; try { header = t.GetRow(0).Row.IsHeader(); } catch (e) {} }
        var start = header ? 1 : 0, rows = [];
        for (var r = start; r < n; r++) {
            var row = t.GetRow(r), cells = [];
            if (row.GetCellsCount() !== cols) return H.fail('irregular_table', 'Row ' + r + ' has merged or missing cells; the table cannot be sorted.');
            for (var c = 0; c < cols; c++) cells.push(row.GetCell(c).GetContent().GetText().replace(/[\r\n]+$/, ''));
            rows.push(cells);
        }
        var key = function (s) { return a.numeric ? parseFloat(String(s).replace(',', '.').replace(/[^0-9.\-]/g, '')) || 0 : String(s).toLowerCase(); };
        var sorted = rows.map(function (x, i) { return { x: x, i: i }; }).sort(function (p, q) {
            var u = key(p.x[a.column]), v = key(q.x[a.column]);
            var d = a.numeric ? u - v : String(u).localeCompare(String(v));
            if (a.descending) d = -d;
            return d || p.i - q.i;
        }).map(function (o) { return o.x; });
        var changed = 0;
        for (var k = 0; k < sorted.length; k++) {
            for (var j = 0; j < cols; j++) {
                if (sorted[k][j] === rows[k][j]) continue;
                var content = t.GetCell(start + k, j).GetContent();
                content.GetRange().Delete();
                content.GetElement(0).AddText(sorted[k][j]);
                changed++;
            }
        }
        return { ok: true, sorted_rows: sorted.length, changed_cells: changed, header_kept: header, tracked: true, done: { column: a.column + 1 } };
    };

    CMD.table_set_cells = function (a, H) {
        var t = H.el(a.table);
        if (H.kind(t) !== 'table') return H.fail('not_a_table', 'Block ' + a.table + ' is not a table.');
        var n = 0;
        for (var k = 0; k < a.cells.length; k++) {
            var it = a.cells[k], cell = t.GetCell(it.row, it.column);
            if (!cell) return H.fail('invalid_argument', 'Cell (' + it.row + ', ' + it.column + ') does not exist.', { field: 'cells[' + k + ']', applied_before_error: n });
            var content = cell.GetContent();
            if (content.GetText().replace(/[\s\r\n]/g, '').length) content.GetRange().Delete();
            content.GetElement(0).AddText(it.text);
            n++;
        }
        return { ok: true, cells: n, tracked: true, done: { n: n } };
    };

    // ---- pictures ----

    CMD.insert_image = function (a, H) {
        var p = Api.CreateParagraph();
        var img = Api.CreateImage(a.url, Math.round(a.width_mm * H.MM_EMU), Math.round(a.height_mm * H.MM_EMU));
        p.AddDrawing(img);
        var wrap = a.wrap || 'inline';
        if (wrap !== 'inline') {
            img.SetWrappingStyle(wrap);
            if (a.align) img.SetHorAlign('margin', a.align);
            img.SetVerPosition('paragraph', 0);
        } else if (a.align) {
            p.SetJc(a.align);
        }
        H.doc.AddElement(a.after + 1, p);
        return { ok: true, block: a.after + 1, width_mm: a.width_mm, height_mm: H.round(a.height_mm, 1), wrap: wrap, tracked: true, done: {} };
    };

    CMD.set_image_layout = function (a, H) {
        var el = H.el(a.block);
        if (H.kind(el) !== 'paragraph') return H.fail('not_a_paragraph', 'Block ' + a.block + ' is not a paragraph.');
        var all = el.GetAllDrawingObjects(), d = all[a.drawing || 0];
        if (!d) return H.fail('no_drawing', 'Block ' + a.block + ' has ' + all.length + ' drawing(s); drawing ' + (a.drawing || 0) + ' does not exist.');
        if (a.width_mm || a.height_mm) {
            var w = d.GetWidth(), h = d.GetHeight();
            var nw = a.width_mm ? a.width_mm * H.MM_EMU : w * (a.height_mm * H.MM_EMU / h);
            var nh = a.height_mm ? a.height_mm * H.MM_EMU : h * (a.width_mm * H.MM_EMU / w);
            if (a.width_mm && a.height_mm && a.keep_aspect !== false) nh = h * nw / w;
            d.SetSize(Math.round(nw), Math.round(nh));
        }
        if (a.wrap) d.SetWrappingStyle(a.wrap);
        var inline = H.wrapOf(d) === 'inline';
        if (a.align) {
            if (inline) el.SetJc(a.align);
            else d.SetHorAlign('margin', a.align);
        }
        if (a.wrap && a.wrap !== 'inline' && !a.align && !inline) d.SetVerPosition('paragraph', 0);
        return { ok: true, wrap: H.wrapOf(d), halign: H.halignOf(d), width_mm: H.round(d.GetWidth() / H.MM_EMU, 1), height_mm: H.round(d.GetHeight() / H.MM_EMU, 1), tracked: false, done: {} };
    };

    // Shrinks the drawings and tables wider than their text column (after a page size change)
    CMD.fit_content = function (a, H) {
        var r = H.fit();
        return { ok: true, fitted: r.n, details: r.list, tracked: false, done: { n: r.n } };
    };
    // ---- page layout ----

    CMD.set_page_setup = function (a, H) {
        var secs = H.doc.GetSections(), idx = [];
        if (a.section === undefined || a.section === 'all') for (var i = 0; i < secs.length; i++) idx.push(i);
        else if (a.section >= 0 && a.section < secs.length) idx.push(a.section);
        else return H.fail('invalid_argument', 'Section ' + a.section + ' does not exist (' + secs.length + ' sections).', { field: 'section' });
        var SIZES = { A3: [297, 420], A4: [210, 297], A5: [148, 210], A6: [105, 148], Letter: [215.9, 279.4], Legal: [215.9, 355.6] };
        var tw = function (mm) { return Math.round(mm * H.MM_TW); };
        for (var k = 0; k < idx.length; k++) {
            var s = secs[idx[k]], w = s.GetPageWidth(), h = s.GetPageHeight();
            if (a.paper) { var sz = SIZES[a.paper]; var land = w > h; w = tw(land ? sz[1] : sz[0]); h = tw(land ? sz[0] : sz[1]); }
            if (a.orientation) { var lo = a.orientation === 'landscape'; if ((w > h) !== lo) { var x = w; w = h; h = x; } }
            if (a.paper || a.orientation) s.SetPageSize(w, h, h >= w);
            var m = a.margins_mm;
            if (m || a.all_margins_mm !== undefined) {
                var all = a.all_margins_mm;
                var L = m && m.left !== undefined ? m.left : (all !== undefined ? all : s.GetPageMarginLeft() / H.MM_TW);
                var T = m && m.top !== undefined ? m.top : (all !== undefined ? all : s.GetPageMarginTop() / H.MM_TW);
                var R = m && m.right !== undefined ? m.right : (all !== undefined ? all : s.GetPageMarginRight() / H.MM_TW);
                var B = m && m.bottom !== undefined ? m.bottom : (all !== undefined ? all : s.GetPageMarginBottom() / H.MM_TW);
                s.SetPageMargins(tw(L), tw(T), tw(R), tw(B));
            }
        }
        // what no longer fits the new text column is shrunk at once (unless fit_content: false)
        var fitted = a.fit_content === false ? { n: 0, list: [] } : H.fit();
        return { ok: true, sections: H.sections(), fitted: fitted.list, overflowing: a.fit_content === false ? H.overflow() : [], tracked: false, done: {} };
    };

    CMD.set_columns = function (a, H) {
        var total = H.count();
        var space = Math.round((a.spacing_mm || 12.5) * 56.6929); // mm -> twips
        var last = H.el(a.to);
        if (a.from > 0 && H.kind(H.el(a.from - 1)) !== 'paragraph')
            return H.fail('unsupported_layout', 'Block ' + (a.from - 1) + ' (just before the range) is a table; a section break cannot be placed there.');
        if (a.to < total - 1 && H.kind(last) !== 'paragraph')
            return H.fail('unsupported_layout', 'Block ' + a.to + ' (end of the range) is a table; a section break cannot be placed there.');
        var target;
        // Close the previous section just before the range (it keeps the original layout)
        if (a.from > 0) H.doc.CreateSection(H.el(a.from - 1));
        if (a.to < total - 1) {
            target = H.doc.CreateSection(last);
            var next = target.GetNext();
            if (next) next.SetType('continuous'); // the content after the range continues on the same page
        } else {
            target = H.doc.GetFinalSection();
        }
        if (a.from > 0) target.SetType('continuous');
        target.SetEqualColumns(a.count, space);
        return {
            ok: true, columns: a.count, from: a.from, to: a.to, tracked: false,
            note: 'Column layout is a section property: Word cannot record it as a tracked revision. It is part of the single undo step.',
            done: { n: a.to - a.from + 1, count: a.count }
        };
    };

    CMD.insert_break = function (a, H) {
        if (a.kind === 'page' || a.kind === 'column') {
            var p = Api.CreateParagraph();
            if (a.kind === 'page') p.AddPageBreak(); else p.AddColumnBreak();
            H.doc.AddElement(a.after + 1, p);
            return { ok: true, block: a.after + 1, tracked: true, note: 'The break is in a new paragraph, block ' + (a.after + 1) + '; later indices moved by +1.', done: { kind: a.kind } };
        }
        var el = H.el(a.after);
        if (H.kind(el) !== 'paragraph') return H.fail('not_a_paragraph', 'A section break must follow a paragraph; block ' + a.after + ' is a ' + H.kind(el) + '.');
        var sec = H.doc.CreateSection(el);
        if (!sec) return H.fail('unsupported_layout', 'A section break cannot be placed after block ' + a.after + '.');
        var next = sec.GetNext();
        if (next) next.SetType(a.kind === 'section_continuous' ? 'continuous' : 'nextPage');
        return { ok: true, tracked: false, done: { kind: 'section' } };
    };

    CMD.set_header_footer = function (a, H) {
        var secs = H.doc.GetSections(), n = 0;
        for (var i = 0; i < secs.length; i++) {
            if (a.section !== undefined && a.section !== 'all' && a.section !== i) continue;
            var s = secs[i];
            if (i > 0 && a.section === undefined) continue; // later sections inherit (linked to previous)
            var c = a.part === 'header' ? s.GetHeader('default', true) : s.GetFooter('default', true);
            if (!c) continue;
            if (a.replace !== false) c.RemoveAllElements();
            var p = Api.CreateParagraph(), txt = a.text || '';
            var parts = txt.split(/(\{PAGE\}|\{NUMPAGES\})/);
            for (var k = 0; k < parts.length; k++) {
                if (parts[k] === '{PAGE}') p.AddPageNumber();
                else if (parts[k] === '{NUMPAGES}') p.AddPagesCount();
                else if (parts[k]) p.AddText(parts[k]);
            }
            if (a.align) p.SetJc(a.align);
            if (a.replace !== false && c.GetElementsCount() > 0) {
                // RemoveAllElements keeps one empty paragraph: fill it
                c.AddElement(0, p);
                if (c.GetElementsCount() > 1) {
                    var last = c.GetElement(c.GetElementsCount() - 1);
                    if (last && last.GetClassType() === 'paragraph' && !last.GetText().replace(/[\r\n\s]/g, '').length) c.RemoveElement(c.GetElementsCount() - 1);
                }
            } else c.Push(p);
            n++;
        }
        return { ok: true, sections: n, tracked: false, done: { part: a.part } };
    };

    CMD.add_comment = function (a, H) {
        var range;
        if (a.find) {
            var ps = H.paras(a.block, a.to !== undefined ? a.to : a.block);
            for (var i = 0; i < ps.length && !range; i++) {
                var rs = H.search(ps[i].p, a.find, false, false);
                if (rs.length) range = rs[0];
            }
            if (!range) return H.fail('not_found', 'Text "' + a.find + '" not found in block ' + a.block + '.');
        } else range = H.range(a.block, a.to !== undefined ? a.to : a.block);
        var c = range.AddComment(a.text, a.author, 'uid-' + a.author);
        if (!c) return H.fail('comment_failed', 'The comment could not be added.');
        return { ok: true, tracked: false, done: {} };
    };
    /* eslint-enable no-undef */

    // ---------------------------------------------------------------------------------
    // Host side: schemas, validation, execution
    // ---------------------------------------------------------------------------------

    var FROM = { type: 'integer', description: 'Index of the first block (inclusive).' };
    var TO = { type: 'integer', description: 'Index of the last block (inclusive).' };
    var TABLE = { type: 'integer', description: 'Block index of the table.' };
    var ALIGN = { type: 'string', 'enum': ['left', 'center', 'right', 'justify'] };
    var WRAP = { type: 'string', 'enum': ['inline', 'square', 'tight', 'through', 'topAndBottom', 'behind', 'inFront'], description: 'inline = in the line of text; square/tight = text flows around; topAndBottom = alone on its lines.' };

    function tool(name, description, properties, required) {
        return { type: 'function', 'function': { name: name, description: description, parameters: { type: 'object', properties: properties, required: required || [], additionalProperties: false } } };
    }

    var TOOLS = [
        // reading
        tool('read_context', 'Re-read blocks of the working scope (index, type, style, text, list, alignment, drawings, tables). The initial context is already in the conversation; call it after changes that moved indices or to verify a result.', {
            from: { type: 'integer', description: 'Optional first block index.' },
            to: { type: 'integer', description: 'Optional last block index.' }
        }),
        tool('document_info', 'Document facts: pages, words, characters, paragraphs (and those of the selection), outline of headings, page setup of each section, header/footer text, elements wider than the text column, comment count. Use it to answer questions and to verify layout goals (page count, fit).', {}),
        tool('find_text', 'Find which blocks contain a text.', {
            text: { type: 'string' }, match_case: { type: 'boolean' }, whole_word: { type: 'boolean' },
            from: { type: 'integer', description: 'Optional first block (default: scope start).' }, to: { type: 'integer', description: 'Optional last block.' }
        }, ['text']),
        tool('explain_layout', 'Layout facts about one block to explain where it is placed: its page, the previous block page and the space left at the bottom of that page, keep-with-next / keep-lines / page-break-before, drawings with their size and wrap, page text height.', {
            block: { type: 'integer' }
        }, ['block']),
        // text
        tool('insert_paragraphs', 'Insert new paragraphs after block `after` (-1 = at the very start). Tracked insertions.', {
            after: { type: 'integer', description: 'Block after which to insert, or -1.' },
            items: {
                type: 'array', minItems: 1, description: 'Paragraphs to insert, in order.',
                items: { type: 'object', properties: { style: { type: 'string', description: 'Paragraph style name, e.g. "Normal", "Heading 2".' }, text: { type: 'string' }, align: ALIGN }, required: ['text'], additionalProperties: false }
            }
        }, ['after', 'items']),
        tool('replace_text', 'Replace text in ONE paragraph: the occurrence(s) of `find`, or the whole paragraph text when `find` is omitted (use it for rewriting, translating or correcting a paragraph while keeping its style). Tracked.', {
            block: { type: 'integer' },
            find: { type: 'string', description: 'Exact text to replace (case-sensitive). Omit to replace the whole paragraph text.' },
            text: { type: 'string', description: 'Replacement text.' },
            all: { type: 'boolean', description: 'Replace every occurrence of `find` in the paragraph.' }
        }, ['block', 'text']),
        tool('find_replace', 'Find and replace across blocks from..to (default: the whole scope), tables included. Tracked.', {
            find: { type: 'string' }, replace: { type: 'string', description: 'Replacement ("" deletes).' },
            match_case: { type: 'boolean' }, whole_word: { type: 'boolean' },
            from: { type: 'integer' }, to: { type: 'integer' }
        }, ['find', 'replace']),
        tool('rewrite_blocks', 'Replace paragraphs from..to by new paragraphs (a different number is allowed, e.g. to shorten or merge). Old ones become tracked deletions, new ones tracked insertions placed after them.', {
            from: FROM, to: TO,
            items: { type: 'array', minItems: 1, items: { type: 'object', properties: { style: { type: 'string' }, text: { type: 'string' } }, required: ['text'], additionalProperties: false } }
        }, ['from', 'to', 'items']),
        tool('delete_blocks', 'Delete blocks from..to (tracked deletion; the blocks stay visible as deletions until accepted).', { from: FROM, to: TO }, ['from', 'to']),
        // styles and structure
        tool('apply_style', 'Apply a paragraph style to every paragraph of blocks from..to, or exactly to the blocks listed in `blocks` (e.g. only the exercise titles). Tracked.', {
            from: { type: 'integer', description: 'First block (with `to`), or use `blocks`.' }, to: { type: 'integer' },
            blocks: { type: 'array', minItems: 1, items: { type: 'integer' }, description: 'Exact blocks to style (non-contiguous allowed).' },
            style: { type: 'string', description: 'Exact paragraph style name from the available styles.' },
            clear_direct_formatting: { type: 'boolean', description: 'Also remove direct bold/italic/underline/strikeout/caps so the style shows.' }
        }, ['style']),
        tool('apply_styles', 'Apply several paragraph styles in one call, one entry per block range. Use it to restyle a document or to turn title lines into headings. Tracked.', {
            changes: {
                type: 'array', minItems: 1,
                items: {
                    type: 'object', additionalProperties: false, required: ['from', 'to', 'style'],
                    properties: { from: FROM, to: TO, style: { type: 'string' }, clear_direct_formatting: { type: 'boolean' } }
                }
            }
        }, ['changes']),
        tool('modify_style', 'Change the definition of a style (font, size, colour, bold, italic, alignment, spacing): every paragraph using it follows. The Word way to apply a style guide ("charte") or to restyle all headings or all body text at once. Not trackable (style definitions have no revisions).', {
            style: { type: 'string', description: 'Exact style name, e.g. "Heading 1", "Normal", "Title".' },
            font: { type: 'string' }, size_pt: { type: 'number', minimum: 1, maximum: 400 }, color: { type: 'string', description: 'Hex colour, e.g. "#1F3864".' },
            bold: { type: 'boolean' }, italic: { type: 'boolean' }, align: ALIGN,
            space_before_pt: { type: 'number', minimum: 0, maximum: 200 }, space_after_pt: { type: 'number', minimum: 0, maximum: 200 }
        }, ['style']),
        tool('change_heading_level', 'Promote (delta -1) or demote (delta +1) the Heading 1-9 paragraphs of blocks from..to. Tracked.', {
            from: FROM, to: TO, delta: { type: 'integer', minimum: -8, maximum: 8 }
        }, ['from', 'to', 'delta']),
        tool('format_text', 'Character formatting of blocks from..to, or only of the occurrences of `find` in them. Tracked.', {
            from: FROM, to: TO,
            find: { type: 'string', description: 'Only format this text (first occurrence unless all=true).' },
            all: { type: 'boolean' }, match_case: { type: 'boolean' }, whole_word: { type: 'boolean' },
            bold: { type: 'boolean' }, italic: { type: 'boolean' }, underline: { type: 'boolean' }, strikethrough: { type: 'boolean' }, caps: { type: 'boolean' },
            font: { type: 'string', description: 'Font family, e.g. "Arial".' },
            size_pt: { type: 'number', minimum: 1, maximum: 400 },
            color: { type: 'string', description: 'Hex colour, e.g. "#1F3864".' },
            highlight: { type: 'string', 'enum': ['yellow', 'green', 'cyan', 'magenta', 'blue', 'red', 'darkBlue', 'darkCyan', 'darkGreen', 'darkMagenta', 'darkRed', 'darkYellow', 'darkGray', 'lightGray', 'black', 'none'] }
        }, ['from', 'to']),
        tool('format_paragraphs', 'Paragraph formatting of the paragraphs of blocks from..to (tables included): alignment, indents (mm, negative first line = hanging), spacing before/after (pt), line spacing (multiple: 1, 1.15, 1.5, 2), keep with next, keep lines together, page break before. Tracked.', {
            from: FROM, to: TO, align: ALIGN,
            indent_left_mm: { type: 'number' }, indent_right_mm: { type: 'number' }, first_line_mm: { type: 'number' },
            space_before_pt: { type: 'number', minimum: 0, maximum: 200 }, space_after_pt: { type: 'number', minimum: 0, maximum: 200 },
            line_spacing: { type: 'number', minimum: 0.5, maximum: 5 },
            keep_with_next: { type: 'boolean' }, keep_lines: { type: 'boolean' }, page_break_before: { type: 'boolean' },
            paragraphs_only_style: { type: 'string', description: 'Optional: only change paragraphs of this style.' }
        }, ['from', 'to']),
        tool('set_list', 'Make paragraphs a bulleted or numbered list (one new list, numbering restarts at `start`), or remove list formatting (kind "none"): either the non-empty paragraphs of blocks from..to, or exactly the blocks listed in `blocks` (e.g. only the exercise titles). Tracked.', {
            from: { type: 'integer', description: 'First block (with `to`), or use `blocks`.' }, to: { type: 'integer' },
            blocks: { type: 'array', minItems: 1, items: { type: 'integer' }, description: 'Exact blocks to put in the list (non-contiguous allowed).' },
            kind: { type: 'string', 'enum': ['bullet', 'numbered', 'none'] },
            level: { type: 'integer', minimum: 0, maximum: 8 },
            format: { type: 'string', 'enum': ['1.', '1)', 'a.', 'a)', 'A.', 'I.', 'i.'], description: 'Numbering format (numbered lists).' },
            start: { type: 'integer', minimum: 0 },
            skip_empty: { type: 'boolean', description: 'Skip empty paragraphs (default true).' }
        }, ['kind']),
        tool('change_list_level', 'Indent (delta +1) or outdent (delta -1) the list items of blocks from..to. Tracked.', {
            from: FROM, to: TO, delta: { type: 'integer', minimum: -8, maximum: 8 }
        }, ['from', 'to', 'delta']),
        tool('insert_table_of_contents', 'Insert a real table of contents field (built from Heading 1-N styles, with page numbers) before block `before`. Headings must exist first: apply heading styles to the title lines in an earlier step.', {
            before: { type: 'integer', description: 'Block before which the table of contents goes (0 = top of the document).' },
            levels: { type: 'integer', minimum: 1, maximum: 9, description: 'Heading levels to include (default 3).' },
            page_numbers: { type: 'boolean' }
        }, ['before']),
        // tables
        tool('list_to_table', 'Convert list paragraphs from..to into a table placed where the list was. Each paragraph becomes a row, split into cells on `separator` (auto-detected). To get a header row ("avec en-tête"), give `header` as the array of column titles you write (one per column, e.g. ["Name", "Group", "Room"]); "first_item" only when the first list item already is a header line; "none" for no header. Tracked.', {
            from: FROM, to: TO,
            header: { description: 'Array of column titles (adds a header row), "first_item" or "none".', anyOf: [{ type: 'array', items: { type: 'string' } }, { type: 'string', 'enum': ['first_item', 'none'] }, { type: 'boolean' }] },
            separator: { type: 'string', description: 'Cell separator inside an item, e.g. " – ", ";", "\\t". Omit to auto-detect.' },
            style: { type: 'string', description: 'Optional table style name (default "Table Grid").' }
        }, ['from', 'to', 'header']),
        tool('create_table', 'Create a table from explicit rows after block `after` (-1 for the start). Tracked.', {
            after: { type: 'integer' },
            rows: { type: 'array', minItems: 1, items: { type: 'array', items: { type: 'string' } } },
            header: { type: 'boolean', description: 'true if the first row is a header row.' },
            style: { type: 'string', description: 'Optional table style name (default "Table Grid").' }
        }, ['after', 'rows']),
        tool('table_rows', 'Insert rows (before/after row `row`, optionally filled with `cells`) or delete rows (`count` rows from `row`). Rows are 0-based. Tracked.', {
            table: TABLE, action: { type: 'string', 'enum': ['insert', 'delete'] }, row: { type: 'integer', minimum: 0 },
            count: { type: 'integer', minimum: 1, maximum: 50 }, position: { type: 'string', 'enum': ['before', 'after'] },
            cells: { type: 'array', items: { type: 'array', items: { type: 'string' } }, description: 'Texts of the inserted rows.' }
        }, ['table', 'action', 'row']),
        tool('table_add_header_row', 'Add a header row at the top of a table with the given column titles (bold, repeated on each page). Use it for "add a header row" on an existing table. Tracked.', {
            table: TABLE, titles: { type: 'array', minItems: 1, items: { type: 'string' }, description: 'One title per column, in order.' }
        }, ['table', 'titles']),
        tool('table_columns', 'Insert or delete table columns (0-based). Not trackable in Word.', {
            table: TABLE, action: { type: 'string', 'enum': ['insert', 'delete'] }, column: { type: 'integer', minimum: 0 },
            count: { type: 'integer', minimum: 1, maximum: 20 }, position: { type: 'string', 'enum': ['before', 'after'] },
            cells: { type: 'array', items: { type: 'array', items: { type: 'string' } }, description: 'Texts of the inserted columns, one array per row.' }
        }, ['table', 'action', 'column']),
        tool('table_format', 'Table formatting: table style, header_row (marks the EXISTING first row as a repeated bold header; it adds no row), banded rows, first column, width in percent of the text width, alignment, column widths in mm.', {
            table: TABLE, style: { type: 'string' }, header_row: { type: 'boolean' }, banded_rows: { type: 'boolean' }, first_column: { type: 'boolean' },
            width_percent: { type: 'number', minimum: 10, maximum: 100 }, align: { type: 'string', 'enum': ['left', 'center', 'right'] },
            column_widths_mm: { type: 'array', items: { type: 'number', minimum: 3, maximum: 500 } }
        }, ['table']),
        tool('table_merge_cells', 'Merge the rectangle of cells from (from_row, from_column) to (to_row, to_column), 0-based. Not trackable.', {
            table: TABLE, from_row: { type: 'integer', minimum: 0 }, from_column: { type: 'integer', minimum: 0 }, to_row: { type: 'integer', minimum: 0 }, to_column: { type: 'integer', minimum: 0 }
        }, ['table', 'from_row', 'from_column', 'to_row', 'to_column']),
        tool('table_sort', 'Sort the rows of a table on one column (0-based); a header row stays on top. Tracked (cell texts are rewritten).', {
            table: TABLE, column: { type: 'integer', minimum: 0 }, descending: { type: 'boolean' }, numeric: { type: 'boolean' },
            has_header: { type: 'boolean', description: 'Default: the first row is kept on top when it is a header row.' }
        }, ['table', 'column']),
        tool('table_set_cells', 'Set the text of table cells (row, column 0-based). Tracked.', {
            table: TABLE,
            cells: { type: 'array', minItems: 1, items: { type: 'object', additionalProperties: false, required: ['row', 'column', 'text'], properties: { row: { type: 'integer', minimum: 0 }, column: { type: 'integer', minimum: 0 }, text: { type: 'string' } } } }
        }, ['table', 'cells']),
        // pictures
        tool('insert_image', 'Insert a picture from a URL (http/https or data:) given by the user, in a new paragraph after block `after`. Size in mm (the aspect ratio is kept when only the width is given). Cannot generate or search for pictures.', {
            after: { type: 'integer' }, url: { type: 'string' },
            width_mm: { type: 'number', minimum: 5, maximum: 400 }, height_mm: { type: 'number', minimum: 5, maximum: 400 },
            wrap: WRAP, align: { type: 'string', 'enum': ['left', 'center', 'right'] }
        }, ['after', 'url']),
        tool('set_image_layout', 'Change a picture (drawing `drawing` of paragraph `block`, from the context): wrap style, horizontal alignment (inline pictures: the paragraph is aligned), size in mm (aspect kept unless keep_aspect=false). Not trackable.', {
            block: { type: 'integer' }, drawing: { type: 'integer', minimum: 0 }, wrap: WRAP, align: { type: 'string', 'enum': ['left', 'center', 'right'] },
            width_mm: { type: 'number', minimum: 3, maximum: 500 }, height_mm: { type: 'number', minimum: 3, maximum: 500 }, keep_aspect: { type: 'boolean' }
        }, ['block']),
        tool('fit_content', 'Shrink every picture and table wider than its text column so it fits (use after reducing the page size or margins, then verify with document_info). Not trackable.', {}),
        // page layout
        tool('set_page_setup', 'Page setup of all sections (default) or one section: paper size, orientation, margins in mm; pictures and tables that no longer fit the text column are shrunk. Not trackable.', {
            section: { type: 'integer', minimum: 0 }, paper: { type: 'string', 'enum': ['A3', 'A4', 'A5', 'A6', 'Letter', 'Legal'] },
            orientation: { type: 'string', 'enum': ['portrait', 'landscape'] },
            all_margins_mm: { type: 'number', minimum: 0, maximum: 100 },
            margins_mm: { type: 'object', additionalProperties: false, properties: { top: { type: 'number', minimum: 0, maximum: 100 }, bottom: { type: 'number', minimum: 0, maximum: 100 }, left: { type: 'number', minimum: 0, maximum: 100 }, right: { type: 'number', minimum: 0, maximum: 100 } } },
            fit_content: { type: 'boolean', description: 'Shrink the pictures and tables that no longer fit the text column (default true).' }
        }),
        tool('set_columns', 'Lay out blocks from..to in `count` newspaper columns, with continuous section breaks around them. Not trackable in Word.', {
            from: FROM, to: TO,
            count: { type: 'integer', minimum: 1, maximum: 6, description: 'Number of columns (1 restores a single column).' },
            spacing_mm: { type: 'number', minimum: 0, maximum: 100 }
        }, ['from', 'to', 'count']),
        tool('insert_break', 'Insert a break after block `after`: page or column break (new paragraph, tracked), or a section break starting on the next page or continuous (not trackable).', {
            after: { type: 'integer' }, kind: { type: 'string', 'enum': ['page', 'column', 'section_next_page', 'section_continuous'] }
        }, ['after', 'kind']),
        tool('set_header_footer', 'Set the text of the header or footer (all pages). Use {PAGE} and {NUMPAGES} for page numbers, e.g. "Page {PAGE} / {NUMPAGES}" or just "{PAGE}", and {DATE} for today\'s date. Not trackable.', {
            part: { type: 'string', 'enum': ['header', 'footer'] }, text: { type: 'string' }, align: ALIGN,
            replace: { type: 'boolean', description: 'Replace the existing content (default true) or add a line.' }
        }, ['part', 'text']),
        tool('add_comment', 'Add a margin comment on block `block` (to `to`), or on the first occurrence of `find` in it.', {
            block: { type: 'integer' }, to: { type: 'integer' }, find: { type: 'string' }, text: { type: 'string' }
        }, ['block', 'text']),
        // analysis (scaffolded mode)
        tool('get_outline_candidates', 'Pre-digested outline: the lines of blocks from..to (default: scope) that look like titles, with their features (style, bold, effective size, numbering pattern and depth, caps, list) and a suggested level (0 = document title, 1 = Heading 1, 2 = Heading 2). Use it before structuring a document or making a table of contents.', {
            from: { type: 'integer' }, to: { type: 'integer' }
        }),
        tool('get_layout_summary', 'Layout summary: page count, page setup and text width of each section, tables, pictures with size and wrap, large text, elements wider than the text column.', {}),
        tool('find_overflowing_elements', 'Pictures and tables wider than their text column (after a page size or margin change).', {}),
        // planning, verification, answers
        tool('submit_plan', 'FIRST call of every request: your plan as structured data (steps may be reading tools, cannot_do or ask_user). Then execute it.', {
            goal: { type: 'string', description: 'The user\'s goal restated in one sentence, with what must be true at the end (checkable).' },
            preconditions: { type: 'array', items: { type: 'string' }, description: 'Facts to check in the document before acting (e.g. "titles are not styled as headings").' },
            steps: {
                type: 'array', minItems: 1,
                items: { type: 'object', additionalProperties: false, required: ['tool', 'purpose'], properties: { tool: { type: 'string', description: 'Tool name used by this step.' }, purpose: { type: 'string' } } }
            }
        }, ['goal', 'steps']),
        tool('verify_result', 'Report whether the goal of your plan is met, after reading the document state you were given.', {
            goal_met: { type: 'boolean' },
            problems: { type: 'array', items: { type: 'string' }, description: 'What is still wrong (empty when the goal is met).' },
            message: { type: 'string', description: 'When the goal is met: one or two sentences for the user, in their language, only if they must know or do something more (e.g. how to print 2 pages per sheet); else empty.' }
        }, ['goal_met']),
        tool('ask_user', 'Ask ONE short clarifying question when the request is truly ambiguous (two very different results are plausible). Give 2-3 short answers to click. Call it alone.', {
            question: { type: 'string' },
            options: { type: 'array', minItems: 2, items: { type: 'string' } }
        }, ['question', 'options']),
        tool('cannot_do', 'Use when the request cannot be done with these tools (or only partly): give one sentence saying what you cannot do, and the closest action you CAN do as a short instruction the user can click. Call it alone.', {
            message: { type: 'string', description: 'One sentence, in the language of the user\'s request.' },
            suggestion: { type: 'string', description: 'A short instruction for the closest doable action, in the language of the user\'s request, e.g. "Insert a picture placeholder".' }
        }, ['message', 'suggestion'])
    ];

    var READ_ONLY = {
        read_context: 1, document_info: 1, find_text: 1, explain_layout: 1, cannot_do: 1, ask_user: 1,
        get_outline_candidates: 1, get_layout_summary: 1, find_overflowing_elements: 1, submit_plan: 1, verify_result: 1
    };
    // Host-side tools (no editor command)
    var HOST_ONLY = { cannot_do: 1, ask_user: 1, submit_plan: 1, verify_result: 1 };
    // Only in the scaffolded mode: planning, verification and the analysis tools
    var SCAFFOLD_ONLY = { submit_plan: 1, verify_result: 1, get_outline_candidates: 1, get_layout_summary: 1, find_overflowing_elements: 1 };

    // Playbooks for common compound requests (scaffolded mode)
    var RECIPES = [
        'Table of contents from unstyled titles: get_outline_candidates -> apply_styles (Heading 1/2/3 from the suggested levels, checked against the text; clear_direct_formatting: true; the document title gets Title) -> insert_table_of_contents (before the first heading, or after the title/cursor) -> verify the outline with document_info.',
        'Change paper size (e.g. A4 -> A5) or make it "2xA5"/"2 per sheet": set_page_setup (paper A5, portrait, margins about 15 mm) -> fit_content -> get_layout_summary to verify (no overflowing element, plausible page count) -> in verify_result.message tell the user to print with "2 pages per sheet" in the print dialog to get two A5 pages on one A4 sheet. Never use columns for this. A booklet is the same A5 setup printed in booklet mode.',
        'Fit on one page: document_info (pages) -> reduce spacing (format_paragraphs space_after_pt 2-4, space_before_pt 0-6), then margins (set_page_setup all_margins_mm 15), then if still needed the body size by 1 pt only (format_text size_pt = current size from the context minus 1; never increase a size) -> document_info again until pages = 1.',
        'Apply a style guide / charte (fonts, colours, sizes): modify_style on Title, Heading 1, Heading 2, Heading 3 and Normal (one call each, same turn); then format_text only on runs that still show other direct formatting.',
        'Create a document (worksheet, letter…) on an empty page: one insert_paragraphs call (after -1, or replace the empty first block) whose items carry their styles: Title for the title, Heading 1 for each exercise / section title and for the answer key ("Corrigé"), Normal for the text; keep it to one page unless asked.',
        'Clean up styles: get_outline_candidates -> one apply_styles call covering every block (titles to headings, body to Normal, clear_direct_formatting on fake titles).',
        'Number the exercises: find the exercise title lines (find_text / get_outline_candidates) -> set_list kind numbered on those lines only (if they already start with a number, keep the text) .',
        'Turn lines into a table / list: list_to_table (header = column titles you write, when a header is wanted) or set_list on exactly the selected blocks.',
        'Add a header row to an existing table: table_add_header_row with one title per column.',
        'Two levels / variants of a worksheet: keep the original; insert_paragraphs after the last block: a Heading 1 title naming the version (e.g. "Version facile"), then each exercise again (title line in the same style as the original exercise titles, then the simplified instruction).',
        'Correct mistakes / translate: replace_text per paragraph (whole text, or `find` for one word), never rewrite_blocks; one call per paragraph, all in the same turn.',
        'Picture next to the text ("à droite du texte"): set_image_layout wrap square + align right.'
    ];
    var UNTRACKED = {
        set_columns: 'columns', set_page_setup: 'page', insert_break: 'section', set_header_footer: 'headerFooter',
        set_image_layout: 'picture', fit_content: 'picture', table_columns: 'table', table_merge_cells: 'table', modify_style: 'style'
    };
    // Steps that need checking: the model verifies them (read back) before finishing
    var NEEDS_VERIFY = { set_page_setup: 1, fit_content: 1, insert_table_of_contents: 1, rewrite_blocks: 1, delete_blocks: 1, insert_break: 1, set_columns: 0 };

    function invalid(field, message, extra) {
        var e = { ok: false, error: { code: 'invalid_argument', field: field, message: message } };
        for (var k in extra || {}) e.error[k] = extra[k];
        return e;
    }
    function isInt(v) { return typeof v === 'number' && isFinite(v) && Math.floor(v) === v; }

    // Validate a value against the subset of JSON Schema used above; returns an error or null
    function check(schema, v, field) {
        if (schema.anyOf) {
            for (var i = 0; i < schema.anyOf.length; i++) if (!check(schema.anyOf[i], v, field)) return null;
            return invalid(field, '"' + field + '" has an invalid type. ' + (schema.description || ''));
        }
        var t = schema.type;
        if (t === 'integer' && !isInt(v)) return invalid(field, '"' + field + '" must be an integer.');
        if (t === 'number' && (typeof v !== 'number' || !isFinite(v))) return invalid(field, '"' + field + '" must be a number.');
        if (t === 'string' && typeof v !== 'string') return invalid(field, '"' + field + '" must be a string.');
        if (t === 'boolean' && typeof v !== 'boolean') return invalid(field, '"' + field + '" must be a boolean.');
        if (t === 'array' && !Array.isArray(v)) return invalid(field, '"' + field + '" must be an array.');
        if (t === 'object' && (!v || typeof v !== 'object' || Array.isArray(v))) return invalid(field, '"' + field + '" must be an object.');
        if (schema['enum'] && schema['enum'].indexOf(v) < 0) return invalid(field, '"' + field + '" must be one of: ' + schema['enum'].join(', ') + '.', { allowed: schema['enum'] });
        if (schema.minimum !== undefined && v < schema.minimum) return invalid(field, '"' + field + '" must be >= ' + schema.minimum + '.');
        if (schema.maximum !== undefined && v > schema.maximum) return invalid(field, '"' + field + '" must be <= ' + schema.maximum + '.');
        if (t === 'array') {
            if (schema.minItems && v.length < schema.minItems) return invalid(field, '"' + field + '" needs at least ' + schema.minItems + ' item(s).');
            if (schema.items) for (var j = 0; j < v.length; j++) { var e = check(schema.items, v[j], field + '[' + j + ']'); if (e) return e; }
        }
        if (t === 'object' && schema.properties) {
            for (var r = 0; r < (schema.required || []).length; r++)
                if (v[schema.required[r]] === undefined) return invalid(field + '.' + schema.required[r], 'Missing "' + schema.required[r] + '" in ' + field + '.');
            for (var key in v) {
                if (!schema.properties[key]) return invalid(field + '.' + key, 'Unknown argument "' + key + '" in ' + field + '.', { allowed: Object.keys(schema.properties) });
                var e2 = check(schema.properties[key], v[key], field + '.' + key);
                if (e2) return e2;
            }
        }
        return null;
    }

    // Validate arguments against the schema and the current scope. Returns null when valid.
    function validate(name, args, ctx) {
        var def = TOOLS.filter(function (t) { return t['function'].name === name; })[0];
        if (!def) return { ok: false, error: { code: 'unknown_tool', message: 'Unknown tool "' + name + '".', available: TOOLS.map(function (t) { return t['function'].name; }) } };
        if (!args || typeof args !== 'object' || Array.isArray(args)) return invalid(null, 'Arguments must be a JSON object.');
        var schema = def['function'].parameters;
        for (var i = 0; i < schema.required.length; i++)
            if (args[schema.required[i]] === undefined) return invalid(schema.required[i], 'Missing required argument "' + schema.required[i] + '".');
        for (var key in args) {
            var p = schema.properties[key];
            if (!p) return invalid(key, 'Unknown argument "' + key + '".', { allowed: Object.keys(schema.properties) });
            var e = check(p, args[key], key);
            if (e) return e;
        }
        var lo = ctx.from, hi = ctx.to;
        var inScope = function (field, v) {
            return (v < lo || v > hi) ? invalid(field, 'Block ' + v + ' is outside the working scope (' + lo + '-' + hi + ').', { scope: [lo, hi] }) : null;
        };
        var inDoc = function (field, v) {
            return (v < 0 || v >= ctx.total) ? invalid(field, 'Block ' + v + ' does not exist (the document has ' + ctx.total + ' blocks).') : null;
        };
        var err = null;
        if (args.from !== undefined && args.to !== undefined && args.from > args.to) return invalid('from', '"from" must be <= "to".');
        ['from', 'to', 'block', 'table'].forEach(function (f) { if (args[f] !== undefined) err = err || inScope(f, args[f]); });
        if (name === 'insert_table_of_contents') err = err || inDoc('before', args.before);
        if (args.after !== undefined && args.after !== -1) err = err || inScope('after', args.after);
        if (args.after === -1 && ctx.scope === 'selection' && lo > 0) err = err || invalid('after', 'Inserting before the selection is outside the working scope; use after >= ' + lo + '.');
        if (err) return err;
        var styleOk = function (f, s) { return ctx.styles.indexOf(s) < 0 ? invalid(f, 'Unknown paragraph style "' + s + '".', { available: ctx.styles }) : null; };
        if (name === 'insert_paragraphs' || name === 'rewrite_blocks') {
            for (var j = 0; j < args.items.length; j++)
                if (args.items[j].style !== undefined && (err = styleOk('items[' + j + '].style', args.items[j].style))) return err;
        }
        if (name === 'apply_styles') {
            for (var m = 0; m < args.changes.length; m++) {
                var ch = args.changes[m], f = 'changes[' + m + ']';
                if (ch.from > ch.to) return invalid(f, '"from" must be <= "to".');
                err = inScope(f + '.from', ch.from) || inScope(f + '.to', ch.to) || styleOk(f + '.style', ch.style);
                if (err) return err;
            }
        }
        if (name === 'apply_style') return styleOk('style', args.style);
        if (name === 'format_paragraphs' && args.paragraphs_only_style !== undefined) return styleOk('paragraphs_only_style', args.paragraphs_only_style);
        if (name === 'table_format' && args.style !== undefined && ctx.tableStyles.indexOf(args.style) < 0)
            return invalid('style', 'Unknown table style "' + args.style + '".', { available: ctx.tableStyles });
        if (name === 'modify_style' && ctx.styles.indexOf(args.style) < 0 && ['Title', 'Normal'].indexOf(args.style) < 0)
            return invalid('style', 'Unknown style "' + args.style + '".', { available: ctx.styles });
        if (name === 'modify_style' && args.color !== undefined && !/^#?[0-9a-f]{6}$/i.test(args.color)) return invalid('color', '"color" must be a hex colour like "#1F3864".');
        if (name === 'format_text' && args.color !== undefined && !/^#?[0-9a-f]{6}$/i.test(args.color)) return invalid('color', '"color" must be a hex colour like "#1F3864".');
        if (name === 'insert_image' && !/^(https?:\/\/|data:image\/)/i.test(args.url)) return invalid('url', 'The picture must be an http(s) URL or a data:image URI given by the user.');
        if (name === 'insert_image' && ctx.instruction !== undefined && ctx.instruction.indexOf(args.url) < 0)
            return invalid('url', 'Only a picture URL written by the user can be inserted; you cannot generate, draw or search for pictures. Call cannot_do with the closest action instead.');
        if ((name === 'set_list' || name === 'apply_style') && !args.blocks && (args.from === undefined || args.to === undefined)) return invalid('from', 'Give from and to, or blocks.');
        if ((name === 'set_list' || name === 'apply_style') && args.blocks) { for (var bi = 0; bi < args.blocks.length; bi++) if ((err = inScope('blocks[' + bi + ']', args.blocks[bi]))) return err; }
        if (name === 'table_merge_cells' && (args.from_row > args.to_row || args.from_column > args.to_column)) return invalid('from_row', 'from_row/from_column must be <= to_row/to_column.');
        if (name === 'set_page_setup' && args.paper === undefined && args.orientation === undefined && args.margins_mm === undefined && args.all_margins_mm === undefined)
            return invalid(null, 'Give at least one of paper, orientation, margins_mm, all_margins_mm.');
        return null;
    }

    // Embed the arguments as JSON. The macro sanitizer of safePluginEval rejects any source matching
    // /\bimport\s+/ (even inside a string literal), so that word is escaped in the payload.
    function payload(obj) {
        return JSON.stringify(obj).replace(/import/g, 'imp\\u006frt');
    }

    function source(cmdName, args) {
        return '(function(){var H=(' + HELPERS.toString() + ')();try{return JSON.stringify((' + CMD[cmdName].toString() + ')(' + payload(args) + ',H));}' +
            'catch(e){return JSON.stringify({ok:false,error:{code:"exception",message:String(e&&e.message||e)}});}})()';
    }

    function callCommand(api, cmdName, args) {
        return new Promise(function (resolve) {
            var attempt = function (n) {
                api.callCommand(source(cmdName, args || {}), true, function (ret) {
                    if (typeof ret === 'string') {
                        try { return resolve(JSON.parse(ret)); } catch (e) {}
                    }
                    // undefined: the editor skipped the command (long action in progress) or the sandbox threw
                    if (n > 0) return setTimeout(function () { attempt(n - 1); }, 300);
                    resolve({ ok: false, error: { code: 'editor_busy', message: 'The editor did not run the command.' } });
                });
            };
            attempt(2);
        });
    }

    function callMethod(api, name, params) {
        return new Promise(function (resolve) { api.callMethod(name, params || [], function (r) { resolve(r); }); });
    }

    // Natural size of a picture, loaded by the browser (null if it cannot be loaded in time)
    function imageSize(url) {
        return new Promise(function (resolve) {
            var img = new Image(), done = false;
            var end = function (v) { if (!done) { done = true; resolve(v); } };
            img.onload = function () { end(img.naturalWidth ? { w: img.naturalWidth, h: img.naturalHeight } : null); };
            img.onerror = function () { end(null); };
            setTimeout(function () { end(null); }, 6000);
            img.src = url;
        });
    }

    function today(lang) {
        var d = new Date();
        try { return d.toLocaleDateString(lang || undefined, { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' }) + ' (' + d.toISOString().slice(0, 10) + ')'; }
        catch (e) { return d.toISOString().slice(0, 10); }
    }

    function todayShort(lang) {
        var d = new Date();
        try { return d.toLocaleDateString(lang || undefined, { year: 'numeric', month: 'long', day: 'numeric' }); } catch (e) { return d.toISOString().slice(0, 10); }
    }

    function createToolset(api, opts) {
        var ts = {
            id: 'word',
            assistantName: ASSISTANT_NAME,
            tools: TOOLS,
            toolsFor: function (mode) {
                return mode === 'direct' ? TOOLS.filter(function (t) { return !SCAFFOLD_ONLY[t['function'].name]; }) : TOOLS;
            },
            isMutating: function (name) { return !READ_ONLY[name]; },
            isTerminal: function (name) { return name === 'cannot_do' || name === 'ask_user'; },
            toolNames: TOOLS.map(function (t) { return t['function'].name; }),
            // Document state given to the model for the verification step
            verifyState: function (ctx) {
                return callCommand(api, 'document_info', { selectedText: '' }).then(function (info) {
                    var t = ctx.touched && ctx.touched.to >= 0 ? ctx.touched : { from: ctx.from, to: ctx.to };
                    var from = Math.max(0, t.from - 1), to = Math.min(ctx.total - 1, t.to + 1);
                    return callCommand(api, 'read_context', { scope: ctx.scope, from: from, to: to, maxBlocks: 120, maxText: 160 }).then(function (c) {
                        delete info.ok;
                        return { document: info, blocks: c.blocks };
                    });
                });
            },
            needsVerify: function (name) { return !!NEEDS_VERIFY[name]; },
            // A deterministic note for the user after some actions (shown with the summary)
            userNote: function (name, args, result) {
                if (name === 'set_page_setup' && args && (args.paper === 'A5' || args.paper === 'A6')) return 'n_print2up';
                return null;
            },
            untrackedNote: function (name, args) {
                if (name === 'insert_break' && args && (args.kind === 'page' || args.kind === 'column')) return null;
                return UNTRACKED[name] || null;
            },

            init: function () {
                return callCommand(api, 'init').then(function (r) {
                    if (!r.ok) throw new Error(r.error && r.error.message || 'cannot read the document');
                    var ctx = {
                        scope: r.scope, from: r.from, to: r.to, total: r.total, styles: r.styles, tableStyles: r.tableStyles || [],
                        revisionsBefore: r.revisions, selectedText: r.selectedText || '', image: r.image, cursor: r.cursor, sections: r.sections
                    };
                    return callCommand(api, 'read_context', { scope: ctx.scope, from: ctx.from, to: ctx.to, maxBlocks: MAX_BLOCKS, maxText: MAX_TEXT })
                        .then(function (c) { ctx.context = c; return ctx; });
                });
            },

            systemPrompt: function (ctx, opts) {
                opts = opts || {};
                if (opts.mode === 'direct') return [
                    'You are Elixir AI, the assistant of a word processor compatible with Microsoft Word. You act on the document only through the tools.',
                    'The document is a list of blocks (paragraphs, tables, content controls) with a 0-based index. Working scope: ' + (ctx.scope === 'selection' ? 'the user\'s selection' : 'the whole document') + ', blocks ' + ctx.from + ' to ' + ctx.to + '. Only edit blocks inside it (page setup, header and footer apply to the document).',
                    'Paragraph styles: ' + ctx.styles.join(', ') + '.',
                    'Table styles: ' + ctx.tableStyles.slice(0, 40).join(', ') + '.',
                    'Today: ' + today(opts.lang) + '. Units: millimetres for lengths, points for font size and spacing.',
                    'Changes are tracked changes; the whole action is one undo step. Indices move after insertions: use the tool results or read_context.',
                    'If the request is impossible with these tools, call cannot_do. If it is truly ambiguous, call ask_user. Before the first tool call write one short sentence with your plan. When finished, reply with one short sentence of what you did, in the language of the user\'s request.'
                ].join('\n');
                return [
                    'You are Elixir AI, the assistant of a word processor compatible with Microsoft Word. You act on the document only through the tools.',
                    'The document is a list of blocks (paragraphs, tables, content controls) with a 0-based index. Working scope: ' + (ctx.scope === 'selection' ? 'the user\'s selection' : 'the whole document') + ', blocks ' + ctx.from + ' to ' + ctx.to + '. Only edit blocks inside it (page setup, header and footer apply to the document).',
                    'Paragraph styles: ' + ctx.styles.join(', ') + '.',
                    'Table styles: ' + ctx.tableStyles.slice(0, 40).join(', ') + '.',
                    'Today: ' + today(opts.lang) + '. Units: millimetres for lengths, points for font size and spacing.',
                    'Changes are tracked changes; the whole action is one undo step.',
                    'Word conventions: the document title uses Title (or Heading 1); section titles Heading 1, sub-sections Heading 2, then Heading 3, following their numbering (1., 1.1, 1.1.1) and their formatting (bigger/bolder = higher level); a subtitle/author/date line under the title is not a heading (Subtitle or Normal); body text is Normal. When titles are faked with direct formatting, apply the heading style with clear_direct_formatting: true. "Make it big" = about 20-28 pt; a title is centred with align "center".',
                    'Method: 1) call submit_plan first (goal with its end condition, preconditions to check, ordered steps each mapped to one tool; follow the playbook when one applies); 2) check preconditions with the analysis tools (get_outline_candidates, get_layout_summary, find_overflowing_elements, document_info, explain_layout) rather than guessing from raw text; 3) execute every step; 4) you will then be shown the document state: compare it with the goal, call verify_result, and fix what is still wrong in the same turn. Never claim a result you did not check; never stop at a partial result. Simple requests get a one-step plan.',
                    'Playbooks:\n- ' + RECIPES.join('\n- '),
                    'Indices move after insertions (insert_paragraphs, create_table, list_to_table, rewrite_blocks, insert_break page, insert_table_of_contents): use the tool results or read_context.',
                    'Never invent content the user did not ask for. When you correct or translate, keep each paragraph and its style (replace_text per paragraph).',
                    'If the request is impossible with these tools, call cannot_do alone (one sentence + the closest doable action as suggestion); never answer a change request with text only. For a question about the document, use the tools to get facts and answer in one or two sentences.',
                    'Before the first tool call write one short sentence with your plan. When finished, reply with one short sentence of what you did, in the language of the user\'s request.'
                ].join('\n');
            },

            userMessage: function (instruction, ctx) {
                var extra = '';
                if (ctx.scope === 'selection' && ctx.selectedText) extra += '\nSelected text: ' + JSON.stringify(ctx.selectedText.slice(0, 1500));
                if (ctx.image) extra += '\nSelected picture: drawing ' + ctx.image.drawing + ' of block ' + ctx.image.block + '.';
                if (ctx.cursor !== null && ctx.cursor !== undefined && ctx.scope === 'document') extra += '\nCursor in block ' + ctx.cursor + '.';
                if (ctx.sections) extra += '\nPage setup: ' + JSON.stringify(ctx.sections);
                return 'Instruction: ' + instruction + extra + '\n\nWorking scope (' + ctx.scope + ', blocks ' + ctx.from + '-' + ctx.to + ' of ' + ctx.total + '):\n' + JSON.stringify(ctx.context.blocks);
            },

            // Forgiving argument names that small models often use: block -> from/to
            normalize: function (name, args) {
                var def = TOOLS.filter(function (t) { return t['function'].name === name; })[0];
                if (!def || !args || typeof args !== 'object') return args;
                var pr = def['function'].parameters.properties;
                if (pr.from && pr.to && !pr.block && args.block !== undefined && args.from === undefined) {
                    args.from = args.block; if (args.to === undefined) args.to = args.block; delete args.block;
                }
                if (pr.block && !pr.from && args.block === undefined && args.from !== undefined && (args.to === undefined || args.to === args.from)) {
                    args.block = args.from; delete args.from; delete args.to;
                }
                return args;
            },
            validate: function (name, args, ctx) { return validate(name, args, ctx); },

            execute: function (name, args, ctx) {
                if (name === 'cannot_do') {
                    ctx.cannotDo = { message: args.message, suggestion: args.suggestion || '' };
                    return Promise.resolve({ ok: true });
                }
                if (name === 'ask_user') {
                    ctx.question = { question: args.question, options: args.options };
                    return Promise.resolve({ ok: true });
                }
                if (name === 'submit_plan') {
                    args.steps.forEach(function (st) { st.tool = String(st.tool).replace(/^functions\./, ''); });
                    var names = TOOLS.map(function (t) { return t['function'].name; });
                    var bad = args.steps.filter(function (st) { return names.indexOf(st.tool) < 0 || st.tool === 'submit_plan' || st.tool === 'verify_result'; });
                    if (bad.length) return Promise.resolve({ ok: false, error: { code: 'invalid_plan', message: 'Unknown or invalid tool in steps: ' + bad.map(function (b) { return b.tool; }).join(', ') + '.', available: names.filter(function (n) { return !HOST_ONLY[n]; }) } });
                    ctx.plan = args;
                    return Promise.resolve({ ok: true, steps: args.steps.length, note: 'Plan accepted: execute every step, then the document state will be shown for verification.' });
                }
                if (name === 'verify_result') {
                    ctx.verification = { goal_met: args.goal_met, problems: args.problems || [], message: args.message || '' };
                    return Promise.resolve({ ok: true });
                }
                if (name === 'read_context') {
                    var from = args.from !== undefined ? Math.max(args.from, ctx.from) : ctx.from;
                    var to = args.to !== undefined ? Math.min(args.to, ctx.to) : ctx.to;
                    return callCommand(api, 'read_context', { scope: ctx.scope, from: from, to: to, maxBlocks: MAX_BLOCKS, maxText: MAX_TEXT });
                }
                var a = Object.assign({}, args);
                if (name === 'document_info') a.selectedText = ctx.scope === 'selection' ? ctx.selectedText : '';
                if (name === 'find_text' || name === 'find_replace' || name === 'get_outline_candidates') {
                    if (a.from === undefined) a.from = ctx.from;
                    if (a.to === undefined) a.to = ctx.to;
                }
                if (name === 'add_comment') a.author = ASSISTANT_NAME;
                // {DATE}: today's date, written out in the editor's language
                var dateText = function (t) { return String(t).replace(/\{DATE[^}]*\}/g, todayShort(ElixirAI.getLang && ElixirAI.getLang())); };
                if (name === 'set_header_footer' && a.text) a.text = dateText(a.text);
                if (name === 'insert_paragraphs') a.items = a.items.map(function (it) { return Object.assign({}, it, { text: dateText(it.text) }); });
                var pre = Promise.resolve();
                if (name === 'insert_image') {
                    pre = imageSize(a.url).then(function (sz) {
                        var ratio = sz ? sz.h / sz.w : 0.75;
                        if (!a.width_mm && !a.height_mm) a.width_mm = Math.min(120, sz ? sz.w * 25.4 / 96 : 80);
                        if (!a.width_mm) a.width_mm = a.height_mm / ratio;
                        if (!a.height_mm) a.height_mm = a.width_mm * ratio;
                    });
                }
                return pre.then(function () { return callCommand(api, name, a); }).then(function (r) {
                    if (r && r.ok) {
                        // keep the working scope in sync with inserted blocks
                        var grow = 0;
                        if (name === 'insert_paragraphs' || name === 'rewrite_blocks') grow = args.items.length;
                        if (name === 'create_table' || name === 'list_to_table' || name === 'insert_image') grow = 1;
                        if (name === 'insert_break' && (args.kind === 'page' || args.kind === 'column')) grow = 1;
                        if (name === 'insert_table_of_contents') grow = 1;
                        ctx.to += grow;
                        ctx.total += grow;
                        ctx.touched = ctx.touched || { from: Infinity, to: -1 };
                        var lo = Infinity, hi = -1, push = function (x, y) { if (x === undefined) return; lo = Math.min(lo, x); hi = Math.max(hi, y === undefined ? x : y); };
                        if (name === 'apply_styles') args.changes.forEach(function (c) { push(c.from, c.to); });
                        else if (args.from !== undefined) push(args.from, args.to + grow);
                        else if (args.block !== undefined) push(args.block, args.to !== undefined ? args.to : args.block);
                        else if (args.table !== undefined) push(args.table);
                        else if (args.after !== undefined) push(args.after + 1, args.after + grow);
                        else if (args.before !== undefined) push(args.before, args.before + grow);
                        else push(ctx.from, ctx.to);
                        if (hi >= 0) {
                            ctx.touched.from = Math.min(ctx.touched.from, Math.max(lo, 0));
                            ctx.touched.to = Math.max(ctx.touched.to, hi);
                        }
                    }
                    return r;
                });
            },

            // Open the undo group and switch tracking on (assistant identity)
            begin: function (ctx) {
                return callMethod(api, 'StartAction', ['GroupActions']).then(function () {
                    return callCommand(api, 'begin', { name: ASSISTANT_NAME });
                }).then(function (r) { ctx.groupOpen = true; return r; });
            },

            end: function (ctx, ok) {
                if (!ctx.groupOpen) return Promise.resolve(null);
                return callCommand(api, 'end', {}).then(function (r) {
                    ctx.groupOpen = false;
                    ctx.revisionsAfter = r && r.revisions;
                    return callMethod(api, 'EndAction', ['GroupActions', '', ok ? '' : 'cancel']).then(function () { return r; });
                });
            },

            // Accept the assistant's revisions: everything when the document had none before, else only
            // those in the touched range (the plugin API cannot filter revisions by author).
            accept: function (ctx) {
                var before = ctx.revisionsBefore ? ctx.revisionsBefore.count : 0;
                if (!before) return callMethod(api, 'AcceptReviewChanges', [true]);
                var t = ctx.touched || { from: ctx.from, to: ctx.to };
                return callCommand(api, 'select', { from: t.from, to: t.to }).then(function () {
                    return callMethod(api, 'AcceptReviewChanges', [false]);
                });
            },

            reject: function (ctx) {
                var before = ctx.revisionsBefore ? ctx.revisionsBefore.count : 0;
                if (!before) return callMethod(api, 'RejectReviewChanges', [true]);
                var t = ctx.touched || { from: ctx.from, to: ctx.to };
                return callCommand(api, 'select', { from: t.from, to: t.to }).then(function () {
                    return callMethod(api, 'RejectReviewChanges', [false]);
                });
            },

            undo: function () { return callMethod(api, 'Undo', []); },
            revisions: function () { return callCommand(api, 'revisions', {}); }
        };
        return ts;
    }

    ElixirAI.toolsets.word = { create: createToolset, tools: TOOLS, _commands: CMD, _source: source };
    return ElixirAI.toolsets.word;
});

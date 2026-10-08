/*
 * SPDX-FileCopyrightText: 2026 ElixirOffice contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 *
 * ElixirOffice Insert tab additions (document editor), like Word's:
 *   - Pages > Cover Page: gallery of cover pages inserted at the start of the document in a
 *     content control (tag "elixir-cover-page"), replaced by the next one, Remove Current Cover Page;
 *   - Table > Quick Tables: Calendar, Double Table, Matrix, Tabular List, With Subheads;
 *   - live preview of the table while hovering the Insert Table grid (an outline drawn over the
 *     page at the insertion point, the document is not modified until the click);
 *   - Header and Footer galleries (Blank, Blank (Three Columns), Title, Page X of Y...) above
 *     the existing Edit and Remove commands.
 * Content is written with the builder API (api.callCommand, one undo step). The only hooks are
 * the controller and module declarations in app.js / app_dev.js and an @import in app.less.
 */
define([
    'core'
], function () {
    'use strict';

    var esc = function (s) { return Common.Utils.String.htmlEncode(s || ''); };
    var COVER_TAG = 'elixir-cover-page';

    DE.Controllers.ElixirInsert = Backbone.Controller.extend({

        initialize: function () {
            var me = this;
            Common.NotificationCenter.on('app:ready', function (mode) { _.defer(function () { me.onAppReady(mode); }); });
        },

        onLaunch: function () {},

        onAppReady: function (mode) {
            var main = DE.getController('Main'), tbc = DE.getController('Toolbar');
            if (!mode || !mode.isEdit || !main || !main.api || !tbc || !tbc.toolbar || this.ready) return;
            this.ready = true;
            this.api = main.api;
            this.toolbar = tbc.toolbar;
            this.mode = mode;
            var $ins = this.toolbar.$el.find('.panel[data-tab=ins]');
            if (!$ins.length) return;
            this.renderCover($ins);
            this.addQuickTables();
            this.bindTablePreview();
            this.addHeaderFooterGalleries();
        },

        lock: function () {
            var _set = Common.enumLock;
            return [_set.previewReviewMode, _set.viewFormMode, _set.lostConnect, _set.disableOnStart, _set.docLockView, _set.docLockForms, _set.docLockComments, _set.viewMode, _set.inHeader, _set.inFootnote, _set.richDelLock, _set.plainDelLock, _set.cantPageBreak];
        },

        done: function () { Common.NotificationCenter.trigger('edit:complete', this.toolbar); },

        run: function (source, cb) {
            var me = this;
            if (!this.api || !this.api.callCommand) return;
            // a command started from the callback of another one is lost: let the first one finish
            this.api.callCommand(source, true, function (r) { setTimeout(function () { cb && cb(r); me.done(); }, 0); });
        },

        // ---- Cover Page --------------------------------------------------------------------------
        covers: function () {
            return [
                {key: 'simple', name: this.txtCoverSimple},
                {key: 'banded', name: this.txtCoverBanded},
                {key: 'centered', name: this.txtCoverCentered},
                {key: 'lines', name: this.txtCoverLines}
            ];
        },

        renderCover: function ($ins) {
            var me = this, $pages = $ins.find('.group[data-group=pages]');
            if (!$pages.length) return;
            $pages.prepend('<span class="btn-slot text x-huge" id="elx-ins-cover"></span>');
            var tpl = _.template('<a id="<%= id %>" tabindex="-1" type="menuitem" class="elx-cover-item"><span class="elx-cover-thumb elx-cover-<%= options.key %>"><i></i><b></b><s></s></span><span class="cap"><%- caption %></span></a>');
            var items = _.map(this.covers(), function (c) { return {caption: c.name, key: c.key, value: c.key, template: tpl}; });
            items.push({caption: '--'}, {caption: this.txtRemoveCover, value: 'remove'});
            this.btnCover = new Common.UI.Button({
                cls: 'btn-toolbar x-huge icon-top', iconCls: 'toolbar__icon btn-elx-cover-page', caption: this.txtCoverPage,
                lock: this.lock(), menu: new Common.UI.Menu({cls: 'elx-cover-menu', style: 'width: 300px;', items: items}),
                dataHint: '1', dataHintDirection: 'bottom', dataHintOffset: 'small'
            });
            this.btnCover.render($ins.find('#elx-ins-cover'));
            this.toolbar.lockControls.push(this.btnCover);
            Common.Utils.lockControls(Common.enumLock.disableOnStart, false, {array: [this.btnCover]});
            this.btnCover.menu.on('item:click', function (menu, item) {
                item.value === 'remove' ? me.removeCover() : me.insertCover(item.value);
            });
        },

        coverTexts: function () {
            var info = this.api.asc_getCoreProps ? this.api.asc_getCoreProps() : null,
                title = info && info.asc_getTitle && info.asc_getTitle();
            return {
                title: title || this.txtDocTitle, subtitle: this.txtDocSubtitle,
                author: AscCommon.UserInfoParser.getParsedName(AscCommon.UserInfoParser.getCurrentName()) || this.txtAuthor,
                date: new Date().toLocaleDateString(this.mode && this.mode.lang ? this.mode.lang.replace('_', '-') : undefined, {year: 'numeric', month: 'long', day: 'numeric'}),
                summary: this.txtAbstract
            };
        },

        insertCover: function (key) {
            this.run('(function () {' +
                'var doc = Api.GetDocument(), key = ' + JSON.stringify(key) + ', t = ' + JSON.stringify(this.coverTexts()) + ', TAG = "' + COVER_TAG + '";' +
                // replace an existing cover page
                'var all = doc.GetAllContentControls ? doc.GetAllContentControls() : [];' +
                'for (var i = 0; i < all.length; i++) { if (all[i].GetTag && all[i].GetTag() === TAG && all[i].Delete) all[i].Delete(false); }' +
                'var sdt = Api.CreateBlockLvlSdt(); sdt.SetTag(TAG); sdt.SetAlias("Cover Page");' +
                'var c = sdt.GetContent(); c.RemoveAllElements();' +
                'var para = function (text, size, bold, color, align, before, after, caps) {' +
                '  var p = Api.CreateParagraph(); if (text) { var r = p.AddText(text); r.SetFontSize(size * 2); r.SetBold(!!bold); r.SetCaps(!!caps);' +
                '  if (color === "accent") r.SetColor(Api.ThemeColor("accent1")); else if (color) r.SetColor(parseInt(color.substr(0,2),16), parseInt(color.substr(2,2),16), parseInt(color.substr(4,2),16)); }' +
                '  p.SetJc(align || "left"); p.SetSpacingBefore((before || 0) * 20); p.SetSpacingAfter((after || 0) * 20); c.Push(p); return p; };' +
                'if (key === "banded") {' +
                '  var b = para(t.title, 36, false, "FFFFFF", "left", 0, 0); b.SetShd("clear", Api.ThemeColor("accent1")); b.SetIndLeft(0);' +
                '  var s = para(t.subtitle, 16, false, "FFFFFF", "left", 0, 0); s.SetShd("clear", Api.ThemeColor("accent1"));' +
                '  para("", 11); para(t.summary, 11, false, "595959", "left", 240, 12); para(t.author, 12, true, null, "left", 0, 0); para(t.date, 11, false, "595959", "left", 0, 0);' +
                '} else if (key === "centered") {' +
                '  para("", 11, false, null, "center", 180, 0); para(t.title, 40, false, "accent", "center", 0, 12, true); para(t.subtitle, 16, false, "595959", "center", 0, 48);' +
                '  para(t.author, 12, true, null, "center", 0, 0); para(t.date, 11, false, "595959", "center", 0, 0);' +
                '} else if (key === "lines") {' +
                '  para(t.date, 11, false, "595959", "right", 0, 96); var l = para(t.title, 36, false, null, "left", 0, 6); l.SetBottomBorder("single", 12, 4, 68, 114, 196);' +
                '  para(t.subtitle, 16, false, "accent", "left", 0, 240); para(t.summary, 11, false, "595959", "left", 0, 24); para(t.author, 12, true, null, "left", 0, 0);' +
                '} else {' +
                '  para("", 11, false, null, "left", 120, 0); para(t.title, 36, false, "accent", "left", 0, 6); para(t.subtitle, 16, false, "595959", "left", 0, 192);' +
                '  para(t.author, 12, true, null, "left", 0, 0); para(t.date, 11, false, "595959", "left", 0, 0);' +
                '}' +
                'var last = Api.CreateParagraph(); last.AddPageBreak(); c.Push(last);' +
                'var trim0 = function (c) { if (c.GetElementsCount() > 1) { var e0 = c.GetElement(0); if (e0 && e0.GetText && !e0.GetText().replace(/\\s/g, "")) c.RemoveElement(0); } }; trim0(c);' +
                'doc.AddElement(0, sdt);' +
                'return true; })()');
        },

        removeCover: function () {
            this.run('(function () { var doc = Api.GetDocument(), all = doc.GetAllContentControls ? doc.GetAllContentControls() : [], n = 0;' +
                'for (var i = 0; i < all.length; i++) { if (all[i].GetTag && all[i].GetTag() === "' + COVER_TAG + '" && all[i].Delete) { all[i].Delete(false); n++; } }' +
                'return n; })()');
        },

        // ---- Quick Tables ------------------------------------------------------------------------
        addQuickTables: function () {
            var me = this, menu = this.toolbar.mnuInsertTable;
            if (!menu) return;
            var sub = new Common.UI.Menu({menuAlign: 'tl-tr', items: [
                {caption: this.txtQtCalendar, value: 'calendar'},
                {caption: this.txtQtDouble, value: 'double'},
                {caption: this.txtQtMatrix, value: 'matrix'},
                {caption: this.txtQtTabular, value: 'tabular'},
                {caption: this.txtQtSubheads, value: 'subheads'}
            ]});
            var idx = _.findIndex(menu.items, function (it) { return it.value === 'sse'; });
            var item = new Common.UI.MenuItem({caption: this.txtQuickTables, value: 'elx-quick', menu: sub});
            menu.insertItem(idx < 0 ? menu.items.length : idx + 1, item);
            sub.on('item:click', function (m, it) { me.insertQuickTable(it.value); });
        },

        quickTable: function (key) {
            var d = new Date(), y = d.getFullYear(), m = d.getMonth(), rows, style = 'Table Grid', header = true;
            var lang = this.mode && this.mode.lang ? this.mode.lang.replace('_', '-') : undefined;
            if (key === 'calendar') {
                var days = [], first = new Date(y, m, 1), start = (first.getDay() + 6) % 7, n = new Date(y, m + 1, 0).getDate();
                for (var k = 0; k < 7; k++) days.push(new Date(2024, 0, 1 + k).toLocaleDateString(lang, {weekday: 'short'}));
                rows = [days];
                var cur = [], day = 1;
                for (var i = 0; i < start; i++) cur.push('');
                while (day <= n) { cur.push(String(day++)); if (cur.length === 7) { rows.push(cur); cur = []; } }
                if (cur.length) { while (cur.length < 7) cur.push(''); rows.push(cur); }
                return {title: first.toLocaleDateString(lang, {month: 'long', year: 'numeric'}), rows: rows, style: 'Grid Table 4 - Accent 1', header: true};
            }
            if (key === 'double') {
                return {rows: [[this.txtQtCol1, this.txtQtCol2, '', this.txtQtCol1, this.txtQtCol2],
                    ['A', '1', '', 'F', '6'], ['B', '2', '', 'G', '7'], ['C', '3', '', 'H', '8'], ['D', '4', '', 'I', '9'], ['E', '5', '', 'J', '10']], style: 'Grid Table 4 - Accent 1', header: true};
            }
            if (key === 'matrix') {
                return {rows: [['', this.txtQtPoint + ' A', this.txtQtPoint + ' B', this.txtQtPoint + ' C', this.txtQtPoint + ' D'],
                    [this.txtQtPoint + ' A', '—', '', '', ''], [this.txtQtPoint + ' B', '87', '—', '', ''], [this.txtQtPoint + ' C', '64', '56', '—', ''], [this.txtQtPoint + ' D', '37', '32', '91', '—']], style: 'Grid Table 1 Light', header: true};
            }
            if (key === 'subheads') {
                return {rows: [[this.txtQtItem, this.txtQtNeeded], [this.txtQtSubhead + ' 1', ''], [this.txtQtItem + ' 1', '10'], [this.txtQtItem + ' 2', '20'], [this.txtQtSubhead + ' 2', ''], [this.txtQtItem + ' 3', '15'], [this.txtQtItem + ' 4', '5']], style: 'Grid Table 4 - Accent 1', header: true};
            }
            return {rows: [[this.txtQtItem, this.txtQtNeeded], [this.txtQtItem + ' 1', '1'], [this.txtQtItem + ' 2', '2'], [this.txtQtItem + ' 3', '3'], [this.txtQtItem + ' 4', '4'], [this.txtQtItem + ' 5', '5']], style: 'List Table 3 - Accent 1', header: true};
        },

        insertQuickTable: function (key) {
            var spec = this.quickTable(key);
            this.run('(function () {' +
                'var doc = Api.GetDocument(), spec = ' + JSON.stringify(spec) + ', out = [];' +
                'if (spec.title) { var p = Api.CreateParagraph(); var r = p.AddText(spec.title); r.SetBold(true); r.SetFontSize(28); out.push(p); }' +
                'var t = Api.CreateTable(spec.rows[0].length, spec.rows.length); t.SetWidth("percent", 100);' +
                'var st = doc.GetStyle(spec.style) || doc.GetStyle("Table Grid"); st && t.SetStyle(st);' +
                't.SetTableLook(!!spec.header, false, true, false, true, false);' +
                'for (var i = 0; i < spec.rows.length; i++) for (var j = 0; j < spec.rows[i].length; j++) { if (spec.rows[i][j]) t.GetRow(i).GetCell(j).GetContent().GetElement(0).AddText(spec.rows[i][j]); }' +
                'out.push(t); out.push(Api.CreateParagraph());' +
                'doc.InsertContent(out); return true; })()');
        },

        // ---- live preview of the Insert Table grid ---------------------------------------------
        bindTablePreview: function () {
            var me = this, tb = this.toolbar, menu = tb.mnuInsertTable;
            if (!menu) return;
            this.$ghost = $('<div class="elx-table-ghost" style="display:none;"></div>').appendTo('#id_main');
            // the grid picker is created with the menu: bind to it once it exists
            var bind = function () {
                var picker = tb.mnuTablePicker;
                if (!picker || picker._elxGhost) return;
                picker._elxGhost = true;
                picker.on('change', function (p, cols, rows) { me.showGhost(cols, rows); });
                picker.on('select', function () { me.hideGhost(); });
            };
            bind();
            menu.on('show:after', bind);
            menu.on('hide:after', function () { me.hideGhost(); });
        },

        showGhost: function (cols, rows) {
            if (!this.api || !cols || !rows) return this.hideGhost();
            var a = this.api.asc_ElixirGetInsertionArea && this.api.asc_ElixirGetInsertionArea();
            if (!a) return this.hideGhost();
            var zoom = (this.api.asc_getZoom ? this.api.asc_getZoom() : 100) / 100, rowH = Math.round(19 * zoom),
                w = a.W, x = a.X, y = a.Y, html = [];
            if (!(w > 0)) return this.hideGhost();
            for (var r = 0; r < rows; r++) {
                html.push('<div class="row">');
                for (var c = 0; c < cols; c++) html.push('<div class="cell"></div>');
                html.push('</div>');
            }
            this.$ghost.html(html.join('')).css({left: x, top: y, width: w, height: rows * rowH}).show();
            this.$ghost.find('.row').css('height', rowH);
        },

        hideGhost: function () { this.$ghost && this.$ghost.hide().empty(); },

        // ---- Header and Footer galleries ------------------------------------------------------------
        addHeaderFooterGalleries: function () {
            var me = this, rt = DE.getController('RibbonTabs'), v = rt && rt.view;
            var add = function (btn, isHeader) {
                if (!btn || !btn.menu || !_.isObject(btn.menu)) return;
                var tpl = _.template('<a id="<%= id %>" tabindex="-1" type="menuitem" class="elx-hf-item"><span class="elx-hf-thumb elx-hf-<%= options.key %>"><i></i><i></i><i></i></span><span class="cap"><%- caption %></span></a>');
                var presets = isHeader ? [
                    {key: 'blank', caption: me.txtHfBlank}, {key: 'three', caption: me.txtHfThree}, {key: 'title', caption: me.txtHfTitle}, {key: 'titledate', caption: me.txtHfTitleDate}
                ] : [
                    {key: 'blank', caption: me.txtHfBlank}, {key: 'three', caption: me.txtHfThree}, {key: 'pagexofy', caption: me.txtHfPageXofY}, {key: 'pagenum', caption: me.txtHfPageNumber}
                ];
                btn.menu.insertItem(0, new Common.UI.MenuItem({caption: '--'}));
                _.each(presets.reverse(), function (p) {
                    btn.menu.insertItem(0, new Common.UI.MenuItem({caption: p.caption, key: p.key, value: 'elx-hf-' + p.key, template: tpl}));
                });
                btn.menu.insertItem(0, new Common.UI.MenuItem({caption: me.txtBuiltIn, template: _.template('<div class="elx-dsg-menu-header">' + esc(me.txtBuiltIn) + '</div>'), disabled: true}));
                btn.menu.on('item:click', function (menu, item) {
                    if (item.options.key && /^elx-hf-/.test(item.value)) me.insertHeaderFooter(isHeader, item.options.key);
                });
            };
            add(v && v.btnInsHeader, true);
            add(v && v.btnInsFooter, false);
        },

        insertHeaderFooter: function (isHeader, key) {
            var t = this.coverTexts();
            this.run('(function () {' +
                'var doc = Api.GetDocument(), sect = doc.GetFinalSection(), key = ' + JSON.stringify(key) + ', isHeader = ' + (isHeader ? 'true' : 'false') + ', t = ' + JSON.stringify({title: t.title, date: t.date, type: this.txtTypeHere, page: this.txtPage, of: this.txtOf}) + ';' +
                'var sections = doc.GetSections ? doc.GetSections() : [sect];' +
                'for (var k = 0; k < sections.length; k++) {' +
                '  var hf = isHeader ? sections[k].GetHeader("default", true) : sections[k].GetFooter("default", true);' +
                '  hf.RemoveAllElements(); var p = Api.CreateParagraph(); p.SetStyle(doc.GetStyle(isHeader ? "Header" : "Footer") || doc.GetStyle("Normal"));' +
                '  if (key === "three") { p.AddText(t.type + "\\t" + t.type + "\\t" + t.type); }' +
                '  else if (key === "title") { var r = p.AddText(t.title); r.SetColor(Api.ThemeColor("accent1")); p.SetJc("center"); p.SetBottomBorder("single", 6, 2, 68, 114, 196); }' +
                '  else if (key === "titledate") { p.AddText(t.title + "\\t\\t" + t.date); p.SetBottomBorder("single", 4, 2, 127, 127, 127); }' +
                '  else if (key === "pagexofy") { p.SetJc("center"); p.AddText(t.page + " "); p.AddPageNumber(); p.AddText(" " + t.of + " "); p.AddPagesCount(); }' +
                '  else if (key === "pagenum") { p.SetJc("right"); p.AddPageNumber(); }' +
                '  else { p.AddText(t.type); }' +
                '  hf.Push(p); var trim0 = function (c) { if (c.GetElementsCount() > 1) { var e0 = c.GetElement(0); if (e0 && e0.GetText && !e0.GetText().replace(/\\s/g, "")) c.RemoveElement(0); } }; trim0(hf);' +
                '}' +
                'return true; })()');
        },

        // ---- strings (English; translations in locale/*.json) ------------------------------------
        txtCoverPage: 'Cover Page',
        txtCoverSimple: 'Simple',
        txtCoverBanded: 'Banded',
        txtCoverCentered: 'Centered',
        txtCoverLines: 'Lines',
        txtRemoveCover: 'Remove Current Cover Page',
        txtDocTitle: '[Document title]',
        txtDocSubtitle: '[Document subtitle]',
        txtAuthor: '[Author name]',
        txtAbstract: '[Draw your reader in with an engaging abstract. It is typically a short summary of the document.]',
        txtQuickTables: 'Quick Tables',
        txtQtCalendar: 'Calendar',
        txtQtDouble: 'Double Table',
        txtQtMatrix: 'Matrix',
        txtQtTabular: 'Tabular List',
        txtQtSubheads: 'With Subheads',
        txtQtCol1: 'Name',
        txtQtCol2: 'Value',
        txtQtPoint: 'City',
        txtQtItem: 'Item',
        txtQtNeeded: 'Needed',
        txtQtSubhead: 'Section',
        txtBuiltIn: 'Built-in',
        txtHfBlank: 'Blank',
        txtHfThree: 'Blank (Three Columns)',
        txtHfTitle: 'Title',
        txtHfTitleDate: 'Title and Date',
        txtHfPageXofY: 'Page X of Y',
        txtHfPageNumber: 'Page Number (Right)',
        txtTypeHere: '[Type here]',
        txtPage: 'Page',
        txtOf: 'of'
    });
});

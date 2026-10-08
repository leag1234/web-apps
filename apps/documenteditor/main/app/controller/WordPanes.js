/*
 * SPDX-FileCopyrightText: 2026 ElixirOffice contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 *
 * DE.Controllers.WordPanes: task panes of Word at the right of the document (sdkjs word/Editor/word-views.js).
 *
 *   Shift+F1   Reveal Formatting: font, paragraph and section formatting at the cursor; the headings open the
 *              Font, Paragraph and Page Setup dialogs
 *   Alt+F10    Selection: the objects of the document (front first), click to select, eye to show / hide,
 *              double click to rename, Show All / Hide All, Bring Forward / Send Backward
 *   Shift+F7   Thesaurus: meanings and synonyms of the word at the cursor, from the offline MyThes thesaurus files
 *              of the dictionaries folder (read with HTTP range requests, nothing leaves the server); click a
 *              synonym to look it up, Insert replaces the word
 * One pane at a time; Esc in the pane or its close button closes it.
 */

define([
    'core'
], function () {
    'use strict';

    DE.Controllers.WordPanes = Backbone.Controller.extend(_.extend({
        initialize: function () {
            this._pane = null;
            Common.NotificationCenter.on('app:ready', this.onAppReady.bind(this));
        },

        onAppReady: function (mode) {
            var me = this;
            this.mode = mode;
            this.api = this.getApplication().getController('Viewport').getApi();
            if (!this.api || !this.api.asc_GetRevealFormatting) return;
            var refresh = _.debounce(function () { me.refresh(); }, 150);
            this.api.asc_registerCallback('asc_onFocusObject', refresh);
            this.api.asc_registerCallback('asc_onDocumentContentReady', refresh);
            this.api.asc_registerCallback('asc_onWordViewChanged', refresh);
            document.addEventListener('keydown', _.bind(this.onKey, this), true);
        },

        onKey: function (e) {
            var K = Common.UI.Keys;
            if (e.keyCode === K.F1 && e.shiftKey && !e.ctrlKey && !e.altKey && !e.metaKey) {
                if (Common.Utils.ModalWindow.isVisible()) return;
                this.stop(e);
                this.toggle('reveal');
            } else if (e.keyCode === K.F10 && e.altKey && !e.ctrlKey && !e.shiftKey && !e.metaKey && !Common.Utils.isMac) {
                if (Common.Utils.ModalWindow.isVisible()) return;
                this.stop(e);
                this.toggle('selection');
            } else if (e.keyCode === K.F7 && e.shiftKey && !e.ctrlKey && !e.altKey && !e.metaKey) {
                if (Common.Utils.ModalWindow.isVisible()) return;
                this.stop(e);
                this.open('thesaurus');
            } else if (e.keyCode === K.ESC && this._pane && $(document.activeElement).closest('#elx-task-pane').length && !$(document.activeElement).is('input')) {
                this.stop(e);
                this.close();
            }
        },

        stop: function (e) {
            e.preventDefault();
            e.stopPropagation();
        },

        // ---- pane frame ---------------------------------------------------------------------
        toggle: function (kind) {
            if (this._pane === kind) this.close();
            else this.open(kind);
        },

        open: function (kind) {
            if (!$('#elx-task-pane-style').length) {
                $('<style id="elx-task-pane-style"></style>').text(
                    '#elx-task-pane{position:absolute;right:0;top:0;bottom:0;width:290px;z-index:500;display:flex;flex-direction:column;' +
                    'background:var(--background-toolbar,#fff);color:var(--text-normal,#333);border-left:1px solid var(--border-toolbar,#ccc);box-shadow:-2px 0 6px rgba(0,0,0,.08);font-size:11px;}' +
                    '#elx-task-pane .elx-tp-head{display:flex;align-items:center;padding:8px 10px;font-size:14px;border-bottom:1px solid var(--border-divider,#e0e0e0);}' +
                    '#elx-task-pane .elx-tp-head span{flex:1;}' +
                    '#elx-task-pane .elx-tp-close{border:0;background:none;font-size:16px;cursor:pointer;color:inherit;padding:0 4px;}' +
                    '#elx-task-pane .elx-tp-body{flex:1;overflow:auto;padding:8px 10px;}' +
                    '#elx-task-pane .elx-rf-sample{border:1px solid var(--border-regular-control,#ccc);padding:6px;margin-bottom:8px;min-height:20px;overflow:hidden;white-space:nowrap;text-overflow:ellipsis;}' +
                    '#elx-task-pane .elx-rf-sec{margin:8px 0 2px;font-weight:bold;}' +
                    '#elx-task-pane .elx-rf-sec a{color:var(--text-link,#2a6ebb);cursor:pointer;text-decoration:none;}' +
                    '#elx-task-pane .elx-rf-key{margin:4px 0 0 10px;font-weight:bold;color:var(--text-secondary,#555);}' +
                    '#elx-task-pane .elx-rf-val{margin:0 0 0 22px;}' +
                    '#elx-task-pane .elx-sp-tools{display:flex;gap:6px;margin-bottom:6px;flex-wrap:wrap;}' +
                    '#elx-task-pane .elx-sp-tools button{font-size:11px;padding:2px 6px;}' +
                    '#elx-task-pane .elx-sp-item{display:flex;align-items:center;padding:3px 4px;cursor:pointer;border:1px solid transparent;}' +
                    '#elx-task-pane .elx-sp-item:hover{background:var(--highlight-button-hover,#e0e0e0);}' +
                    '#elx-task-pane .elx-sp-item.selected{background:var(--highlight-button-pressed,#cbcbcb);}' +
                    '#elx-task-pane .elx-sp-item.hidden-obj .elx-sp-name{opacity:.55;}' +
                    '#elx-task-pane .elx-sp-name{flex:1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}' +
                    '#elx-task-pane .elx-sp-eye{border:0;background:none;cursor:pointer;color:inherit;font-size:12px;width:22px;}' +
                    '#elx-task-pane .elx-sp-empty{color:var(--text-secondary,#777);margin-top:8px;}'
                ).appendTo('head');
            }
            this.close(true);
            var me = this, title = {reveal: this.txtReveal, selection: this.txtSelection, thesaurus: this.txtThesaurus}[kind];
            var $host = $('#editor_sdk').parent();
            var $pane = $('<div id="elx-task-pane" role="complementary" tabindex="-1"><div class="elx-tp-head"><span></span>' +
                '<button class="elx-tp-close" type="button">&#x2715;</button></div><div class="elx-tp-body"></div></div>');
            $pane.attr('aria-label', title).attr('data-pane', kind).find('.elx-tp-head span').text(title);
            $pane.find('.elx-tp-close').attr('title', this.txtClose).on('click', function () { me.close(); });
            $host.css('position') === 'static' && $host.css('position', 'relative');
            $host.append($pane);
            this._pane = kind;
            this.refresh();
        },

        close: function (silent) {
            $('#elx-task-pane').remove();
            this._pane = null;
            if (!silent && this.api) {
                this.api.asc_enableKeyEvents(true);
                Common.NotificationCenter.trigger('edit:complete');
            }
        },

        refresh: function () {
            if (!this._pane || !this.api) return;
            var $body = $('#elx-task-pane .elx-tp-body');
            if (!$body.length) return;
            if (this._pane === 'reveal') this.renderReveal($body);
            else if (this._pane === 'selection') this.renderSelection($body);
            else if (!$body.children().length) this.renderThesaurus($body);
        },

        // ---- Reveal Formatting --------------------------------------------------------------
        len: function (mm) {
            var M = Common.Utils.Metric;
            var unit = M.getCurrentMetricName();
            return (Math.round(M.fnRecalcFromMM(mm) * 100) / 100) + (unit === '"' ? '' : ' ') + unit;
        },

        pt: function (mm) {
            return (Math.round(mm * 72 / 25.4 * 10) / 10) + ' pt';
        },

        renderReveal: function ($body) {
            var f = this.api.asc_GetRevealFormatting();
            if (!f) return;
            var me = this, font = f.font || {}, par = f.paragraph || {}, sec = f.section, rows = [];
            var section = function (title, action) { rows.push('<div class="elx-rf-sec"><a data-action="' + action + '">' + _.escape(title) + '</a></div>'); };
            var item = function (key, values) {
                values = _.compact(values);
                if (!values.length) return;
                rows.push('<div class="elx-rf-key">' + _.escape(key) + '</div><div class="elx-rf-val">' + values.map(_.escape).join(', ') + '</div>');
            };

            section(this.txtFont, 'font');
            item(this.txtFontLabel, [font.name ? font.name : this.txtMixed, font.size ? font.size + ' pt' : null, font.bold && this.txtBold, font.italic && this.txtItalic, font.underline && this.txtUnderline]);
            if (font.lang !== null && font.lang !== undefined) {
                var lang = Common.util.LanguageInfo.getLocalLanguageName(font.lang);
                item(this.txtLanguage, [lang && lang[1] ? lang[1] : String(font.lang)]);
            }
            var color = font.color ? (font.color.auto ? this.txtAutomatic : '#' + [font.color.r, font.color.g, font.color.b].map(function (c) { return ('0' + (c || 0).toString(16)).slice(-2); }).join('').toUpperCase()) : null;
            item(this.txtFontColor, [color]);
            item(this.txtEffects, [font.strikeout && this.txtStrikethrough, font.caps && this.txtAllCaps, font.smallCaps && this.txtSmallCaps, font.hidden && this.txtHidden,
                font.position === 1 && this.txtSuperscript, font.position === 2 && this.txtSubscript, font.highlight && this.txtHighlight]);
            font.spacing && item(this.txtCharSpacing, [(font.spacing > 0 ? this.txtExpanded : this.txtCondensed) + ' ' + this.pt(Math.abs(font.spacing))]);

            section(this.txtParagraph, 'paragraph');
            item(this.txtParaStyle, [par.style]);
            var align = {};
            align[AscCommon.align_Left] = this.txtLeft; align[AscCommon.align_Center] = this.txtCentered;
            align[AscCommon.align_Right] = this.txtRight; align[AscCommon.align_Justify] = this.txtJustified;
            item(this.txtAlignment, [par.align !== null ? align[par.align] || this.txtJustified : null]);
            item(this.txtOutlineLevel, [par.outlineLevel !== null && par.outlineLevel !== undefined ? Common.Utils.String.format(this.txtLevel, par.outlineLevel + 1) : this.txtBodyText]);
            var first = par.firstLine || 0;
            item(this.txtIndentation, [par.left !== null ? this.txtIndLeft + ' ' + this.len(par.left) : null, par.right !== null ? this.txtIndRight + ' ' + this.len(par.right) : null,
                first > 0.01 ? this.txtFirstLine + ' ' + this.len(first) : first < -0.01 ? this.txtHanging + ' ' + this.len(-first) : null]);
            var line = null;
            if (par.line !== null && par.lineRule !== null) {
                if (par.lineRule === Asc.linerule_Auto) line = Math.abs(par.line - 1) < 0.01 ? this.txtSingle : Math.abs(par.line - 2) < 0.01 ? this.txtDouble : Common.Utils.String.format(this.txtMultiple, Math.round(par.line * 100) / 100);
                else line = (par.lineRule === Asc.linerule_Exact ? this.txtExactly : this.txtAtLeast) + ' ' + this.pt(par.line);
            }
            item(this.txtSpacing, [par.before !== null ? this.txtBefore + ' ' + this.pt(par.before) : null, par.after !== null ? this.txtAfter + ' ' + this.pt(par.after) : null, line && this.txtLineSpacing + ' ' + line]);
            item(this.txtPagination, [par.widowControl && this.txtWidow, par.keepNext && this.txtKeepNext, par.keepLines && this.txtKeepLines, par.pageBreakBefore && this.txtBreakBefore]);
            par.numbered && item(this.txtBullets, [this.txtNumbered]);

            if (sec) {
                section(this.txtSection, 'section');
                item(this.txtMargins, [this.txtIndLeft + ' ' + this.len(sec.left), this.txtIndRight + ' ' + this.len(sec.right), this.txtTop + ' ' + this.len(sec.top), this.txtBottom + ' ' + this.len(sec.bottom)]);
                sec.columns > 1 && item(this.txtLayout, [Common.Utils.String.format(this.txtColumns, sec.columns)]);
                item(this.txtPaper, [this.txtWidth + ' ' + this.len(sec.width), this.txtHeight + ' ' + this.len(sec.height)]);
            }

            var sample = (this.api.asc_GetSelectedText ? this.api.asc_GetSelectedText() : '') || this.txtSampleText;
            $body.html('<div class="elx-rf-key" style="margin-left:0">' + _.escape(this.txtSelectedText) + '</div><div class="elx-rf-sample"></div>' +
                '<div class="elx-rf-key" style="margin-left:0">' + _.escape(this.txtFormatting) + '</div>' + rows.join(''));
            $body.find('.elx-rf-sample').text(String(sample).replace(/[\r\n]+/g, ' ').slice(0, 80)).css({'font-family': font.name || '', 'font-weight': font.bold ? 'bold' : '', 'font-style': font.italic ? 'italic' : ''});
            $body.find('.elx-rf-sec a').on('click', function () { me.openDialog($(this).attr('data-action')); });
        },

        openDialog: function (action) {
            var app = this.getApplication(), fmt = app.getController('WordFormatting'), tb = app.getController('Toolbar');
            if (!this.mode || !this.mode.isEdit) return;
            if (action === 'font') fmt && fmt.openFontDialog && fmt.openFontDialog();
            else if (action === 'paragraph') fmt && fmt.openParagraphDialog && fmt.openParagraphDialog();
            else if (action === 'section') tb && tb.onPageMarginsSelect && tb.onPageMarginsSelect(null, {value: 'advanced'});
        },

        // ---- Selection pane -----------------------------------------------------------------
        objectName: function (o) {
            if (o.name) return o.name;
            var base = {picture: this.txtPicture, shape: this.txtShape, textbox: this.txtTextBox, group: this.txtGroup, chart: this.txtChart, smartart: this.txtSmartArt}[o.kind] || this.txtObject;
            return base + ' ' + o.number;
        },

        renderSelection: function ($body) {
            var me = this, list = this.api.asc_GetSelectionPaneObjects() || [], edit = !!(this.mode && this.mode.isEdit);
            var $tools = $('<div class="elx-sp-tools"></div>');
            if (edit) {
                $('<button type="button" class="btn btn-text-default" data-cmd="show"></button>').text(this.txtShowAll).appendTo($tools);
                $('<button type="button" class="btn btn-text-default" data-cmd="hide"></button>').text(this.txtHideAll).appendTo($tools);
                $('<button type="button" class="btn btn-text-default" data-cmd="forward"></button>').text(this.txtForward).attr('title', this.txtForward).appendTo($tools);
                $('<button type="button" class="btn btn-text-default" data-cmd="backward"></button>').text(this.txtBackward).attr('title', this.txtBackward).appendTo($tools);
            }
            var $list = $('<div class="elx-sp-list" role="listbox"></div>');
            _.each(list, function (o) {
                var $it = $('<div class="elx-sp-item" role="option"><span class="elx-sp-name"></span><button type="button" class="elx-sp-eye"></button></div>');
                $it.attr('data-id', o.id).toggleClass('selected', !!o.selected).toggleClass('hidden-obj', !!o.hidden).attr('aria-selected', !!o.selected);
                $it.find('.elx-sp-name').text(me.objectName(o));
                // open eye: shown; a dash: hidden (Word draws a crossed eye)
                $it.find('.elx-sp-eye').html(o.hidden ? '&#x2014;' : '&#x1F441;').attr('title', o.hidden ? me.txtShow : me.txtHide).attr('aria-pressed', !o.hidden).prop('disabled', !edit);
                $list.append($it);
            });
            $body.empty().append($tools, $list);
            !list.length && $body.append($('<div class="elx-sp-empty"></div>').text(this.txtNoObjects));

            $list.find('.elx-sp-item').on('click', function (e) {
                if ($(e.target).closest('.elx-sp-eye').length) return;
                me.api.asc_SelectPaneObject($(this).attr('data-id'));
                me.refresh();
            }).on('dblclick', function () {
                if (!edit) return;
                var id = $(this).attr('data-id'), $name = $(this).find('.elx-sp-name'), old = $name.text();
                var $input = $('<input type="text" class="form-control" style="width:100%;height:20px;font-size:11px;">').val(old);
                $name.empty().append($input);
                me.api.asc_enableKeyEvents(false);
                $input.focus().select();
                var done = function (save) {
                    var v = $input.val().trim();
                    $input.off();
                    save && v && v !== old && me.api.asc_RenamePaneObject(id, v);
                    me.api.asc_enableKeyEvents(true);
                    me.refresh();
                };
                $input.on('keydown', function (e) {
                    if (e.keyCode === Common.UI.Keys.RETURN) { me.stop(e); done(true); }
                    else if (e.keyCode === Common.UI.Keys.ESC) { me.stop(e); done(false); }
                }).on('blur', function () { done(true); });
            });
            $list.find('.elx-sp-eye').on('click', function () {
                var $it = $(this).closest('.elx-sp-item');
                me.api.asc_SetPaneObjectsHidden($it.attr('data-id'), !$it.hasClass('hidden-obj'));
                me.refresh();
            });
            $tools.find('button').on('click', function () {
                var cmd = $(this).attr('data-cmd');
                if (cmd === 'show' || cmd === 'hide') me.api.asc_SetPaneObjectsHidden(null, cmd === 'hide');
                else {
                    // the selected floating object, as Arrange > Bring Forward / Send Backward
                    var props = new Asc.asc_CImgProperty();
                    props.put_ChangeLevel(cmd === 'forward' ? Asc.c_oAscChangeLevel.BringForward : Asc.c_oAscChangeLevel.BringBackward);
                    me.api.ImgApply(props);
                }
                me.refresh();
            });
        },

        // ---- Thesaurus ----------------------------------------------------------------------
        // MyThes files: <dir>.idx is "word|offset" per line, <dir>.dat holds "word|n" then n lines "(pos)|syn|syn..."
        thesaurusFiles: function (lcid) {
            var info = Common.util.LanguageInfo.getLocalLanguageName(lcid), code = info && info[0] ? info[0].replace('-', '_') : 'en_US',
                dirs = [code];
            /^en_/.test(code) && code !== 'en_US' && dirs.push('en_US');
            var list = [];
            _.each(dirs, function (d) { list.push(d + '/' + d + '_thes', d + '/th_' + d + '_v2', d + '/th_' + d); });
            return list;
        },

        loadThesaurus: function (lcid) {
            var me = this;
            this._thes = this._thes || {};
            var key = String(lcid);
            if (this._thes[key]) return this._thes[key];
            var base = '../../../../dictionaries/', files = this.thesaurusFiles(lcid);
            var tryNext = function (i) {
                if (i >= files.length) return Promise.resolve(null);
                return fetch(base + files[i] + '.idx').then(function (r) {
                    if (!r.ok) throw new Error('missing');
                    return r.arrayBuffer();
                }).then(function (buf) {
                    var head = new TextDecoder('latin1').decode(new Uint8Array(buf, 0, Math.min(40, buf.byteLength))).split(/\r?\n/)[0].replace(/^\uFEFF|^\u00EF\u00BB\u00BF/, '').trim();
                    var enc = /utf-?8/i.test(head) ? 'utf-8' : 'latin1';
                    var lines = new TextDecoder(enc).decode(buf).split(/\r?\n/), index = new Map();
                    for (var n = 2; n < lines.length; n++) {
                        var bar = lines[n].lastIndexOf('|');
                        bar > 0 && index.set(lines[n].substring(0, bar).toLowerCase(), +lines[n].substring(bar + 1));
                    }
                    return {dat: base + files[i] + '.dat', enc: enc, index: index};
                }).catch(function () { return tryNext(i + 1); });
            };
            this._thes[key] = tryNext(0);
            return this._thes[key];
        },

        lookup: function (thes, word) {
            var key = String(word || '').toLowerCase(), offset = thes && thes.index.get(key);
            if (offset === undefined) return Promise.resolve(null);
            // the offsets of some files do not count the byte order mark of the .dat: read a little before
            var from = Math.max(0, offset - 8);
            return fetch(thes.dat, {headers: {Range: 'bytes=' + from + '-' + (offset + 16383)}}).then(function (r) { return r.arrayBuffer(); }).then(function (buf) {
                var lines = new TextDecoder(thes.enc).decode(buf).split(/\r?\n/), start = 0;
                while (start < lines.length - 1 && lines[start].toLowerCase().indexOf(key + '|') !== 0) start++;
                lines = lines.slice(start);
                var first = lines[0].split('|'), count = +first[1] || 0, meanings = [];
                for (var i = 1; i <= count && i < lines.length; i++) {
                    var parts = lines[i].split('|');
                    meanings.push({pos: (parts[0] || '').replace(/[()]/g, ''), words: parts.slice(1).filter(Boolean)});
                }
                return meanings;
            });
        },

        renderThesaurus: function ($body, word) {
            var me = this, lcid = (this.api.asc_GetRevealFormatting() || {font: {}}).font.lang || 1033;
            if (word === undefined) word = (this.api.asc_GetSelectedText && this.api.asc_GetSelectedText(true) || '').trim() || (this.api.asc_GetCurrentWord && this.api.asc_GetCurrentWord()) || '';
            word = String(word).replace(/[\r\n]+/g, ' ').trim().split(/\s+/).slice(0, 3).join(' ');
            this._thesWord = this._thesWord || word;
            $body.html('<div style="display:flex;gap:4px;margin-bottom:8px;"><input type="text" class="form-control elx-th-input" style="flex:1;height:22px;"><button type="button" class="btn btn-text-default elx-th-go"></button></div><div class="elx-th-result"></div>');
            var $input = $body.find('.elx-th-input').val(word).attr('aria-label', this.txtSearchFor), $res = $body.find('.elx-th-result');
            $body.find('.elx-th-go').text(this.txtLookUp);
            var show = function (w) {
                $input.val(w);
                $res.text(me.txtSearching);
                me.loadThesaurus(lcid).then(function (thes) {
                    if (!thes) { $res.text(me.txtNoThesaurus); return null; }
                    return me.lookup(thes, w).then(function (meanings) {
                        $res.empty();
                        if (!meanings || !meanings.length) { $res.text(Common.Utils.String.format(me.txtNoResults, w)); return; }
                        _.each(meanings, function (m) {
                            $('<div class="elx-rf-sec"></div>').text(m.pos ? m.pos : w).appendTo($res);
                            _.each(m.words, function (syn) {
                                var $row = $('<div class="elx-sp-item elx-th-syn"><span class="elx-sp-name"></span><button type="button" class="elx-sp-eye elx-th-insert"></button></div>');
                                $row.find('.elx-sp-name').text(syn);
                                $row.find('.elx-th-insert').text('\u21B5').attr('title', me.txtInsert);
                                $row.on('click', function (e) {
                                    if ($(e.target).closest('.elx-th-insert').length) me.insertSynonym(syn);
                                    else show(syn.replace(/\s*\(.*\)$/, ''));
                                });
                                $res.append($row);
                            });
                        });
                    });
                });
            };
            $input.on('focus', function () { me.api.asc_enableKeyEvents(false); }).on('blur', function () { me.api.asc_enableKeyEvents(true); })
                .on('keydown', function (e) { if (e.keyCode === Common.UI.Keys.RETURN) { me.stop(e); show($input.val().trim()); } });
            $body.find('.elx-th-go').on('click', function () { show($input.val().trim()); });
            word ? show(word) : $res.text(this.txtTypeWord);
        },

        insertSynonym: function (syn) {
            if (!this.mode || !this.mode.isEdit || !this.api) return;
            var text = syn.replace(/\s*\(.*\)$/, '');
            var selected = (this.api.asc_GetSelectedText && this.api.asc_GetSelectedText(true) || '').trim();
            // the selection, or the word of the cursor
            if (selected) this.api.asc_PasteData(AscCommon.c_oAscClipboardDataFormat.Text, text);
            else this.api.asc_ReplaceCurrentWord(0, text);
            this.api.asc_enableKeyEvents(true);
            Common.NotificationCenter.trigger('edit:complete');
        },

        txtThesaurus: 'Thesaurus',
        txtSearchFor: 'Search for',
        txtLookUp: 'Look Up',
        txtSearching: 'Searching...',
        txtNoThesaurus: 'There is no thesaurus for the language of this text.',
        txtNoResults: 'No results for "{0}".',
        txtTypeWord: 'Type a word and press Enter.',
        txtInsert: 'Insert',
        txtReveal: 'Reveal Formatting',
        txtSelection: 'Selection',
        txtClose: 'Close',
        txtSelectedText: 'Selected text',
        txtSampleText: 'Sample Text',
        txtFormatting: 'Formatting of selected text',
        txtFont: 'Font',
        txtFontLabel: 'FONT:',
        txtMixed: '(mixed)',
        txtBold: 'Bold',
        txtItalic: 'Italic',
        txtUnderline: 'Underline',
        txtLanguage: 'LANGUAGE:',
        txtFontColor: 'FONT COLOR:',
        txtAutomatic: 'Automatic',
        txtEffects: 'EFFECTS:',
        txtStrikethrough: 'Strikethrough',
        txtAllCaps: 'All caps',
        txtSmallCaps: 'Small caps',
        txtHidden: 'Hidden',
        txtSuperscript: 'Superscript',
        txtSubscript: 'Subscript',
        txtHighlight: 'Highlight',
        txtCharSpacing: 'CHARACTER SPACING:',
        txtExpanded: 'Expanded by',
        txtCondensed: 'Condensed by',
        txtParagraph: 'Paragraph',
        txtParaStyle: 'PARAGRAPH STYLE:',
        txtAlignment: 'ALIGNMENT:',
        txtLeft: 'Left',
        txtCentered: 'Centered',
        txtRight: 'Right',
        txtJustified: 'Justified',
        txtOutlineLevel: 'OUTLINE LEVEL:',
        txtLevel: 'Level {0}',
        txtBodyText: 'Body Text',
        txtIndentation: 'INDENTATION:',
        txtIndLeft: 'Left:',
        txtIndRight: 'Right:',
        txtFirstLine: 'First line:',
        txtHanging: 'Hanging:',
        txtSpacing: 'SPACING:',
        txtBefore: 'Before:',
        txtAfter: 'After:',
        txtLineSpacing: 'Line spacing:',
        txtSingle: 'single',
        txtDouble: 'double',
        txtMultiple: 'Multiple {0} li',
        txtExactly: 'Exactly',
        txtAtLeast: 'At least',
        txtPagination: 'PAGINATION:',
        txtWidow: 'Widow/Orphan control',
        txtKeepNext: 'Keep with next',
        txtKeepLines: 'Keep lines together',
        txtBreakBefore: 'Page break before',
        txtBullets: 'BULLETS AND NUMBERING:',
        txtNumbered: 'Numbered or bulleted list',
        txtSection: 'Section',
        txtMargins: 'MARGINS:',
        txtTop: 'Top:',
        txtBottom: 'Bottom:',
        txtLayout: 'LAYOUT:',
        txtColumns: '{0} columns',
        txtPaper: 'PAPER:',
        txtWidth: 'Width:',
        txtHeight: 'Height:',
        txtShowAll: 'Show All',
        txtHideAll: 'Hide All',
        txtForward: 'Bring Forward',
        txtBackward: 'Send Backward',
        txtShow: 'Show',
        txtHide: 'Hide',
        txtNoObjects: 'There are no objects in this document.',
        txtPicture: 'Picture',
        txtShape: 'Shape',
        txtTextBox: 'Text Box',
        txtGroup: 'Group',
        txtChart: 'Chart',
        txtSmartArt: 'Diagram',
        txtObject: 'Object'
    }, DE.Controllers.WordPanes || {}));
});

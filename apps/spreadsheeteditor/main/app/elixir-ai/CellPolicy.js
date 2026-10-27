/*
 * SPDX-FileCopyrightText: 2026 ElixirOffice contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 */
define([], function () {
    'use strict';
    // Pure, closure-free factory: the SAME checks run before dispatch and inside callCommand.
    return function CellPolicy() {
        var P = {}, MAX = 10000;
        P.col = function (s) {
            var n = 0; for (var i = 0; i < s.length; i++) n = n * 26 + s.charCodeAt(i) - 64;
            return n;
        };
        P.letter = function (n) {
            var s = ''; while (n > 0) { n--; s = String.fromCharCode(65 + n % 26) + s; n = Math.floor(n / 26); } return s;
        };
        P.address = function (r1, c1, r2, c2) {
            return P.letter(c1) + r1 + (r1 === r2 && c1 === c2 ? '' : ':' + P.letter(c2) + r2);
        };
        P.range = function (s, max) {
            if (typeof s !== 'string' || s.length > 40) throw Error('Use one bounded A1 range on the current worksheet.');
            var m = /^\$?([A-Za-z]{1,3})\$?([1-9][0-9]{0,6})(?::\$?([A-Za-z]{1,3})\$?([1-9][0-9]{0,6}))?$/.exec(s);
            if (!m) throw Error('Use one bounded A1 range on the current worksheet.');
            var c1 = P.col(m[1].toUpperCase()), r1 = +m[2], c2 = P.col((m[3] || m[1]).toUpperCase()), r2 = +(m[4] || m[2]);
            var n = (r2 - r1 + 1) * (c2 - c1 + 1);
            if (c2 > 16384 || r2 > 1048576 || c1 > c2 || r1 > r2 || n > (max || MAX)) throw Error('Range exceeds the bounded action limits.');
            return {r1:r1, c1:c1, r2:r2, c2:c2, rows:r2-r1+1, cols:c2-c1+1, cells:n, address:P.address(r1,c1,r2,c2)};
        };
        P.contains = function (outer, inner) {
            return inner.r1 >= outer.r1 && inner.c1 >= outer.c1 && inner.r2 <= outer.r2 && inner.c2 <= outer.c2;
        };
        var functions = ('SUM AVERAGE MIN MAX COUNT COUNTA COUNTBLANK COUNTIF COUNTIFS SUMIF SUMIFS AVERAGEIF AVERAGEIFS IF IFS IFERROR IFNA AND OR NOT ROUND ROUNDUP ROUNDDOWN ABS INT MOD CEILING FLOOR PRODUCT SUMPRODUCT MEDIAN STDEV STDEV.S STDEV.P VAR VAR.S VAR.P VLOOKUP HLOOKUP XLOOKUP INDEX MATCH CONCAT CONCATENATE TEXTJOIN LEFT RIGHT MID LEN TRIM CLEAN UPPER LOWER PROPER SUBSTITUTE REPLACE FIND SEARCH TEXT VALUE DATE YEAR MONTH DAY WEEKDAY TODAY NOW TRUE FALSE ISBLANK ISNUMBER ISTEXT ISERROR ISNA').split(' ');
        P.formula = function (f) {
            if (typeof f !== 'string' || f.length > 4096 || f.charAt(0) !== '=' || /[\u0000-\u001f\u007f]/.test(f)) throw Error('Use a bounded spreadsheet formula.');
            var outside = '', quoted = false;
            for (var i = 1; i < f.length; i++) {
                var ch = f.charAt(i);
                if (ch === '"') { if (quoted && f.charAt(i+1) === '"') { i++; continue; } quoted = !quoted; outside += ' '; }
                else if (!quoted) outside += ch;
            }
            if (quoted || /[^A-Za-z0-9_$.,:()+*/^&%<>=! \-]/.test(outside) || outside.indexOf('!') >= 0 || /\/\*|\/\//.test(outside)) throw Error('External, dynamic and unsupported formulas are refused.');
            outside = outside.replace(/\b([A-Za-z][A-Za-z0-9.]*)\s*(?=\()/g, function (all, fn) {
                if (functions.indexOf(fn.toUpperCase()) < 0) throw Error('Unsupported formula function: ' + fn);
                return ' ';
            });
            var referenced = 0;
            outside = outside.replace(/\$?[A-Za-z]{1,3}\$?[1-9][0-9]*(?::\$?[A-Za-z]{1,3}\$?[1-9][0-9]*)?/g, function (ref) {
                referenced += P.range(ref).cells; if (referenced > MAX) throw Error('Formula references exceed the action limit.'); return ' ';
            });
            outside = outside.replace(/\b(?:TRUE|FALSE)\b/gi, ' ').replace(/(?:[0-9]+(?:\.[0-9]*)?|\.[0-9]+)(?:[Ee][+-]?[0-9]+)?/g, ' ');
            if (/[A-Za-z_$:]/.test(outside)) throw Error('Named ranges and whole columns are unsupported.');
            if (!outside.trim() && f.trim() === '=') throw Error('Empty formula.');
            return f;
        };
        P.canonicalFormula=function(f) { return String(f).split(/("(?:[^"]|"")*")/g).map(function(part,i){return i%2?part:part.replace(/\s+/g,'').toUpperCase();}).join(''); };
        P.functions=functions.slice();
        P.localFormula=function(formula,locale) {
            P.formula(formula);
            if(!locale || typeof locale.decimal!=='string' || locale.decimal.length!==1 || [',',';'].indexOf(locale.separator)<0 || !locale.functions || typeof locale.functions!=='object')throw Error('Public formula locale metadata is unavailable.');
            return formula.split(/("(?:[^"]|"")*")/g).map(function(part,i){
                if(i%2)return part;
                part=part.replace(/\b([A-Za-z][A-Za-z0-9.]*)\s*(?=\()/g,function(all,fn){return locale.functions[fn.toUpperCase()]||fn;});
                // Rewrite argument commas before numeric decimal dots; quoted literals stay exact.
                part=part.replace(/,/g,locale.separator);
                return part.replace(/(?:[0-9]+(?:\.[0-9]*)?|\.[0-9]+)(?:[Ee][+-]?[0-9]+)?/g,function(n){return n.replace('.',locale.decimal);});
            }).join('');
        };
        P.references = function(f) {
            P.formula(f);
            var unquoted=f.replace(/"(?:[^"]|"")*"/g,' ');
            unquoted=unquoted.replace(/\b[A-Za-z][A-Za-z0-9.]*\s*(?=\()/g,' ');
            return (unquoted.match(/\$?[A-Za-z]{1,3}\$?[1-9][0-9]*(?::\$?[A-Za-z]{1,3}\$?[1-9][0-9]*)?/g)||[]).map(function(ref){return P.range(ref);});
        };
        P.translate = function (f, dr, dc) {
            P.formula(f);
            var chunks = f.split(/("(?:[^"]|"")*")/g);
            f = chunks.map(function (s, i) {
                if (i % 2) return s;
                return s.replace(/(\$?)([A-Za-z]{1,3})(\$?)([1-9][0-9]*)(?![A-Za-z0-9_(])/g, function (all, ca, col, ra, row) {
                    var c = P.col(col.toUpperCase()) + (ca ? 0 : dc), r = +row + (ra ? 0 : dr);
                    if (c < 1 || c > 16384 || r < 1 || r > 1048576) throw Error('Translated formula leaves the worksheet.');
                    return ca + P.letter(c) + ra + r;
                });
            }).join('');
            return P.formula(f);
        };
        P.number = function (value, decimal, group) {
            if (typeof value !== 'string' || ['.', ','].indexOf(decimal) < 0 || ['', '.', ',', ' ', '\u00a0', '\u202f'].indexOf(group) < 0 || decimal === group) throw Error('Choose distinct decimal and grouping separators.');
            var s = value.trim(); if (!s) return null;
            var esc = function (x) { return x.replace(/[.*+?^$(){}|[\]\\]/g, '\\$&'); };
            var integer = group ? '(?:[0-9]+|[0-9]{1,3}(?:' + esc(group) + '[0-9]{3})+)' : '[0-9]+';
            if (!(new RegExp('^[+-]?' + integer + '(?:' + esc(decimal) + '[0-9]+)?$')).test(s)) throw Error('Ambiguous or invalid number. Requested grouping '+(group?'U+'+group.charCodeAt(0).toString(16).toUpperCase().padStart(4,'0'):'none')+'; observed whitespace '+Array.from(new Set((s.match(/\s/g)||[]).map(function(ch){return 'U+'+ch.charCodeAt(0).toString(16).toUpperCase().padStart(4,'0');}))).join(', ')+'. Use the exact separator shown by read_range; U+0020 is an ordinary space.');
            if (group) s = s.split(group).join('');
            var n = Number(s.replace(decimal, '.')); if (!isFinite(n)) throw Error('Non-finite number.'); return n;
        };
        P.date = function (value, format) {
            if (typeof value !== 'string') throw Error('Date must be text.');
            var m, y, month, d;
            if (format === 'YYYY-MM-DD' && (m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(value.trim()))) { y=+m[1]; month=+m[2]; d=+m[3]; }
            else if ((format === 'DD/MM/YYYY' || format === 'MM/DD/YYYY') && (m=/^(\d{2})\/(\d{2})\/(\d{4})$/.exec(value.trim()))) {
                y=+m[3]; month=+(format === 'DD/MM/YYYY' ? m[2] : m[1]); d=+(format === 'DD/MM/YYYY' ? m[1] : m[2]);
            } else throw Error('Date requires an explicit supported input format.');
            var leap = y % 4 === 0 && (y % 100 !== 0 || y % 400 === 0);
            if (y < 1900 || y > 9999 || month < 1 || month > 12 || d < 1 || d > [31,leap?29:28,31,30,31,30,31,31,30,31,30,31][month-1]) throw Error('Invalid calendar date.');
            return [y,month,d];
        };
        P.clean = function (cells, a) {
            var rows = cells.map(function (row) { return row.map(function (c) { return Object.assign({},c); }); });
            if (a.operation === 'dedupe') {
                var seen = Object.create(null), out = [];
                rows.forEach(function (row,i) {
                    if (row.some(function(c){return !!c.formula;})) throw Error('Dedupe of formulas is refused to preserve references.');
                    var key=JSON.stringify(row.map(function(c){return [c.kind || 'text',c.value];}));
                    if (i === 0 && a.header) out.push(row);
                    else if (!seen[key]) { seen[key]=true; out.push(row); }
                });
                return out;
            }
            if (a.operation === 'split') {
                if (typeof a.delimiter !== 'string' || !a.delimiter || a.delimiter.length > 20 || rows.some(function(row){return row.length!==1 || row[0].formula;})) throw Error('Split requires one literal column and an explicit delimiter.');
                var width = 0;
                rows = rows.map(function(row) {
                    var parts=String(row[0].value).split(a.delimiter).map(function(v){return {value:v,formula:null};}); width=Math.max(width,parts.length); return parts;
                });
                if (width > 32 || width * rows.length > MAX) throw Error('Split output exceeds limits.');
                rows.forEach(function(row){while(row.length<width)row.push({value:'',formula:null});});
                return rows;
            }
            rows.forEach(function (row) { row.forEach(function (c) {
                if (c.formula || c.value === '') return;
                if (a.operation === 'trim') c.value = String(c.value).trim().replace(/[ \t\u00a0\u202f]+/g,' ');
                else if (a.operation === 'numbers') { c.number=P.number(String(c.value),a.decimal,a.group); c.kind='number'; }
                else if (a.operation === 'dates') { c.date=P.date(String(c.value),a.format); c.kind='date'; }
                else throw Error('Unsupported cleaning operation.');
            }); });
            return rows;
        };
        P.pivotEstimate=function(cells,a) {
            var rows=Object.create(null),cols=Object.create(null);
            cells.slice(1).forEach(function(row){
                rows[JSON.stringify(a.rows.map(function(i){return [row[i-1].kind,row[i-1].value];}))]=true;
                cols[JSON.stringify((a.columns||[]).map(function(i){return [row[i-1].kind,row[i-1].value];}))]=true;
            });
            var size=(Object.keys(rows).length+a.rows.length+3)*(Math.max(1,Object.keys(cols).length)*a.values.length+a.rows.length+1);
            if(size>2000)throw Error('Pivot output exceeds the2000-cell safety limit; select fewer groups or fields.');
            return size;
        };
        P.summary = function (cells) {
            var out = {rows:cells.length, columns:[]}, width=cells.length?cells[0].length:0;
            for (var c=0;c<width;c++) {
                var numbers=[], blank=0;
                cells.forEach(function(row){var x=row[c];if(x.value==='')blank++;else if(x.kind==='number' && isFinite(+x.value))numbers.push(+x.value);});
                var sum=numbers.reduce(function(a,b){return a+b;},0);
                out.columns.push({column:c+1,numeric:numbers.length,sum:sum,min:numbers.length?Math.min.apply(null,numbers):null,max:numbers.length?Math.max.apply(null,numbers):null,mean:numbers.length?sum/numbers.length:null,blank:blank});
            }
            return out;
        };
        var fields = {
            read_range:['range'], explain_formula:['range'], summarize_range:['range'],
            write_formula:['range','formula','fill'],
            clean_range:['range','operation','delimiter','target','format','decimal','group','header'],
            sort_range:['range','header','keys'],
            filter_range:['range','column','operator','value'],
            create_chart:['range','type','title','anchor'],
            create_pivot:['range','rows','columns','values','name'],
            submit_plan:['goal','steps','read_only'], verify_result:['goal_met','problems','message'],
            ask_user:['question','options'], cannot_do:['message','suggestion']
        };
        P.names = Object.keys(fields);
        P.mutating = function(n){return ['write_formula','clean_range','sort_range','filter_range','create_chart','create_pivot'].indexOf(n)>=0;};
        var str = function(x,max){return typeof x==='string' && x.length>0 && x.length<=(max||4096);};
        var integer = function(n,min,max){return typeof n==='number' && isFinite(n) && Math.floor(n)===n && n>=min && n<=max;};
        var keys = function(o,allowed){return o && typeof o==='object' && !Array.isArray(o) && Object.keys(o).every(function(k){return allowed.indexOf(k)>=0;});};
        P.validate = function (n,a) {
            try {
                if (!fields[n] || !keys(a,fields[n])) throw Error('Unknown tool or argument.');
                var r;
                if (fields[n].indexOf('range')>=0) r=P.range(a.range);
                if (n==='write_formula') {
                    P.formula(a.formula); if(a.fill!==undefined && typeof a.fill!=='boolean')throw Error('fill must be boolean.');
                    if(a.fill===false && r.cells!==1)throw Error('Use fill:true to translate a formula over a range.');
                    for(var dr=0;dr<r.rows;dr++)for(var dc=0;dc<r.cols;dc++)P.translate(a.formula,dr,dc);
                } else if (n==='clean_range') {
                    if(['trim','split','dedupe','dates','numbers'].indexOf(a.operation)<0)throw Error('Unknown cleaning operation.');
                    if(a.operation==='split' && (!str(a.delimiter,20) || !a.target || P.range(a.target).cols>32 || r.cols!==1))throw Error('Split needs a delimiter, single source column and bounded target.');
                    if(a.operation==='dates')P.date('2026-01-02',a.format==='YYYY-MM-DD'?a.format:'YYYY-MM-DD');
                    if(a.operation==='dates' && ['YYYY-MM-DD','DD/MM/YYYY','MM/DD/YYYY'].indexOf(a.format)<0)throw Error('Specify the date format.');
                    if(a.operation==='numbers')P.number('1',a.decimal,a.group);
                    if(a.header!==undefined && typeof a.header!=='boolean')throw Error('header must be boolean.');
                    if(a.target!==undefined)P.range(a.target);
                } else if(n==='sort_range') {
                    if(typeof a.header!=='boolean' || !Array.isArray(a.keys) || a.keys.length<1 || a.keys.length>3)throw Error('Sort needs 1-3 keys and an explicit header flag.');
                    a.keys.forEach(function(k){if(!keys(k,['column','order']) || !integer(k.column,1,r.cols) || ['asc','desc'].indexOf(k.order)<0)throw Error('Invalid sort key.');});
                } else if(n==='filter_range') {
                    if(r.rows<2 || !integer(a.column,1,r.cols) || ['equals','not_equals','contains','greater','less','at_least','at_most'].indexOf(a.operator)<0 || !(typeof a.value==='number'&&isFinite(a.value) || str(a.value,200)))throw Error('Invalid filter criterion.');
                } else if(n==='create_chart') {
                    if(r.rows<2 || r.cols<2 || ['bar','barStacked','lineNormal','pie','scatter'].indexOf(a.type)<0 || !str(a.title,200))throw Error('Invalid chart source, type or title.');
                    if(a.anchor!==undefined && P.range(a.anchor).cells!==1)throw Error('Anchor must be one cell.');
                } else if(n==='create_pivot') {
                    if(r.rows<2 || !Array.isArray(a.rows) || a.rows.length<1 || a.rows.length>4 || !Array.isArray(a.values) || a.values.length<1 || a.values.length>4)throw Error('Pivot needs row and value fields.');
                    (a.rows.concat(a.columns || [])).forEach(function(c){if(!integer(c,1,r.cols))throw Error('Invalid pivot field.');});
                    if(a.columns!==undefined && (!Array.isArray(a.columns)||a.columns.length>4))throw Error('Invalid column fields.');
                    a.values.forEach(function(v){if(!keys(v,['column','aggregation']) || !integer(v.column,1,r.cols) || ['Sum','Count','Average','Min','Max'].indexOf(v.aggregation)<0)throw Error('Invalid pivot aggregation.');});
                    if(!str(a.name,31) || /[:\\/?*\[\]]/.test(a.name))throw Error('Choose a valid new worksheet name.');
                } else if(n==='submit_plan') {
                    if(typeof a.read_only!=='boolean' || !str(a.goal,1000) || !Array.isArray(a.steps) || a.steps.length>12)throw Error('Invalid plan.');
                    a.steps.forEach(function(s){if(!keys(s,['tool','purpose'])||P.names.indexOf(s.tool)<0||s.tool==='submit_plan'||!str(s.purpose,500))throw Error('Each plan step must contain ONLY tool and purpose; put parameters in the later tool call. Use a known tool name.');if(a.read_only&&P.mutating(s.tool))throw Error('A read-only plan cannot contain editing steps. If the user explicitly requests changes, submit read_only:false with those editing steps; otherwise use only read tools.');});
                } else if(n==='verify_result') {
                    if(typeof a.goal_met!=='boolean' || a.problems!==undefined&&(!Array.isArray(a.problems)||a.problems.length>12||a.problems.some(function(p){return !str(p,500);})) || a.message!==undefined&&!str(a.message,2000))throw Error('Invalid verification.');
                } else if(n==='ask_user') {
                    if(!str(a.question,1000) || !Array.isArray(a.options) || a.options.length<1 || a.options.length>5 || a.options.some(function(x){return !str(x,300);}))throw Error('Invalid question.');
                } else if(n==='cannot_do' && (!str(a.message,1000) || a.suggestion!==undefined&&!str(a.suggestion,1000)))throw Error('Invalid refusal.');
                return null;
            } catch(e) { return {ok:false,error:{code:'invalid_arguments',message:e.message}}; }
        };
        return P;
    };
});

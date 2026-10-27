/*
 * SPDX-FileCopyrightText: 2026 ElixirOffice contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 */
define([], function () {
    'use strict';
    // Closure-free: serialize this same policy into the static builder command.
    return function () {
        var names=['read_slides','create_slides','rewrite_text','set_layout','apply_theme','align_shapes','set_notes','summarize_deck','submit_plan','verify_result','ask_user','cannot_do'];
        var edits=['create_slides','rewrite_text','set_layout','apply_theme','align_shapes','set_notes'];
        var modes=['left','center','right','top','middle','bottom','horizontal','vertical'];
        function fail(message){throw Error(message);}
        function object(a,keys,required) {
            if(!a || typeof a!=='object' || Array.isArray(a))fail('Expected an argument object.');
            Object.keys(a).forEach(function(k){if(keys.indexOf(k)<0)fail('Unknown argument: '+k);});
            (required||keys).forEach(function(k){if(!Object.prototype.hasOwnProperty.call(a,k))fail('Missing argument: '+k);});
        }
        function text(s,max,empty) {
            if(typeof s!=='string'||s.length>max||(!empty&&!s.trim())||/[\u0000-\u0008\u000b\u000c\u000e-\u001f\u007f]/.test(s))fail('Text is empty, invalid or exceeds the limit.');
            return s;
        }
        function number(n,max){if(!Number.isSafeInteger(n)||n<1||n>max)fail('Expected a bounded positive integer.');return n;}
        function list(a,max,min){if(!Array.isArray(a)||a.length>(max||10)||a.length<(min===undefined?1:min))fail('Invalid list size.');return a;}
        function unique(a){if(a.some(function(v,i){return a.indexOf(v)!==i;}))fail('Duplicate targets are not allowed.');}
        function slides(a){list(a,10);a.forEach(function(n){number(n,100);});unique(a);}
        function id(s){text(s,100);if(!/^[a-zA-Z0-9_-]+$/.test(s))fail('Invalid public shape identifier.');}
        function mutating(n){return edits.indexOf(n)>=0;}
        function validate(n,a) {
            try {
                if(names.indexOf(n)<0)fail('Unknown presentation tool.');
                if(n==='read_slides'){object(a,['slides']);slides(a.slides);}
                else if(n==='summarize_deck'){object(a,[]);}
                else if(n==='create_slides') {
                    object(a,['slides']);list(a.slides,10);var chars=0;
                    a.slides.forEach(function(s){
                        object(s,['title','bullets','notes'],['title','bullets']);text(s.title,160);chars+=s.title.length;
                        list(s.bullets,8,0).forEach(function(v){text(v,1000);chars+=v.length;});
                        if(s.notes!==undefined){text(s.notes,8000,true);chars+=s.notes.length;}
                    });
                    if(chars>40000)fail('New slide text exceeds 40000 characters.');
                } else if(n==='rewrite_text') {
                    object(a,['slide','shape_id','paragraphs']);number(a.slide,100);id(a.shape_id);list(a.paragraphs,50);
                    var indexes=[],chars=0;
                    a.paragraphs.forEach(function(p){object(p,['index','expected','text','max_chars'],['index','expected','text']);number(p.index,50);text(p.expected,4096,true);text(p.text,4096,true);if(/[\r\n]/.test(p.text))fail('Rewrite one existing paragraph at a time; do not add line breaks.');if(p.max_chars!==undefined){number(p.max_chars,4096);var length=Array.from(p.text).length;if(length>p.max_chars)fail('Replacement has '+length+' characters, exceeding max_chars '+p.max_chars+'. Shorten it before writing.');}indexes.push(p.index);chars+=p.text.length;});
                    unique(indexes);if(chars>20000)fail('Rewrite text exceeds the limit.');
                } else if(n==='set_notes') {
                    object(a,['slide','expected','text','mode']);number(a.slide,100);text(a.expected,8000,true);text(a.text,8000,true);
                    if(['append','replace'].indexOf(a.mode)<0)fail('Choose append or replace notes explicitly.');
                    if(a.mode==='append'&&a.expected&&(a.text===a.expected||a.text.indexOf(a.expected+'\n')===0))fail('Append text must contain ONLY new note paragraphs, without the existing notes. Use replace for complete desired notes.');
                } else if(n==='set_layout'||n==='apply_theme') {
                    object(a,n==='set_layout'?['slides','master','layout']:['slides','master']);slides(a.slides);number(a.master,32);
                    if(n==='set_layout')number(a.layout,100);
                } else if(n==='align_shapes') {
                    object(a,['slide','shape_ids','alignment']);number(a.slide,100);list(a.shape_ids,30,2);a.shape_ids.forEach(id);unique(a.shape_ids);
                    if(modes.indexOf(a.alignment)<0)fail('Unsupported alignment.');
                    if(['horizontal','vertical'].indexOf(a.alignment)>=0&&a.shape_ids.length<3)fail('Distribution needs at least three shapes.');
                } else if(n==='submit_plan') {
                    object(a,['goal','read_only','steps']);text(a.goal,1000);if(typeof a.read_only!=='boolean')fail('An explicit read_only flag is required.');
                    list(a.steps,12,0).forEach(function(s){
                        object(s,['tool','purpose']);if(names.indexOf(s.tool)<0||['submit_plan','verify_result','ask_user','cannot_do'].indexOf(s.tool)>=0)fail('Use a supported action or read tool in the plan.');
                        text(s.purpose,500);if(a.read_only&&mutating(s.tool))fail('A read-only plan cannot include editing tools.');
                    });
                } else if(n==='verify_result') {
                    object(a,['goal_met','problems','message'],['goal_met']);if(typeof a.goal_met!=='boolean')fail('goal_met must be boolean.');
                    if(a.problems!==undefined)list(a.problems,20,0).forEach(function(s){text(s,1000);});
                    if(a.message!==undefined)text(a.message,4000,true);
                } else if(n==='ask_user') {
                    object(a,['question','options']);text(a.question,1000);list(a.options,5).forEach(function(s){text(s,500);});
                } else if(n==='cannot_do') {
                    object(a,['message','suggestion'],['message']);text(a.message,2000);if(a.suggestion!==undefined)text(a.suggestion,2000,true);
                }
                return null;
            } catch(e){return {ok:false,error:{code:'invalid_arguments',message:String(e.message||e)}};}
        }
        function paragraphs(raw) {
            var body=raw&&raw.content&&raw.content.content;
            if(!body)return [];
            if(body.type!=='drawingDocContent'||!Array.isArray(body.content)||body.content.length>50)fail('Unsupported or oversized text body.');
            return body.content.map(function(p,i){
                if(p.type!=='paragraph'||!Array.isArray(p.content)||p.content.length>200)fail('Unsupported paragraph content.');
                var value='',plain=!(p.changes&&p.changes.length),style=null;
                p.content.forEach(function(run){
                    if(['run','endRun'].indexOf(run.type)<0||run.reviewType&&run.reviewType!=='common'||run.footnotes&&run.footnotes.length||run.endnotes&&run.endnotes.length)plain=false;
                    if(!Array.isArray(run.content)){plain=false;return;}
                    run.content.forEach(function(v){if(typeof v==='string')value+=v;else plain=false;});
                    if(run.type==='run'&&run.content.length){
                        var current=JSON.stringify(run.rPr||{});
                        if(style!==null&&style!==current)plain=false;
                        style=current;
                    }
                });
                text(value,4096,true);
                return {index:i+1,text:value,characters:Array.from(value).length,plain:plain};
            });
        }
        function notes(raw) {
            var tree=raw&&raw.notes&&raw.notes.cSld&&raw.notes.cSld.spTree;
            if(!tree)return {exists:false,text:'',paragraphs:[]};
            if(!Array.isArray(tree)||tree.length>20)fail('Unsupported notes page.');
            var body=tree.filter(function(s){var ph=s.nvSpPr&&s.nvSpPr.nvPr&&s.nvSpPr.nvPr.ph;return ph&&ph.type==='body';});
            if(body.length>1)fail('Ambiguous notes body.');
            if(!body.length||!body[0].content)return {exists:false,text:'',paragraphs:[]};
            var rows=paragraphs(body[0]),value=rows.map(function(p){return p.text;}).join('\n');
            text(value,8000,true);
            return {exists:true,text:value,paragraphs:rows};
        }
        function align(boxes,mode,width,height) {
            list(boxes,30,2);if(modes.indexOf(mode)<0)fail('Unsupported alignment.');
            if(!Number.isFinite(width)||!Number.isFinite(height)||width<=0||height<=0)fail('Invalid slide size.');
            var out=boxes.map(function(b){
                if(b.locked)fail('A selected drawing is locked.');
                if([b.x,b.y,b.w,b.h].some(function(n){return !Number.isFinite(n);})||b.x<0||b.y<0||b.w<=0||b.h<=0||b.x+b.w>width||b.y+b.h>height)fail('A drawing is outside the slide or has invalid dimensions.');
                return {id:b.id,x:b.x,y:b.y,w:b.w,h:b.h};
            });
            var left=Math.min.apply(null,out.map(function(b){return b.x;})),right=Math.max.apply(null,out.map(function(b){return b.x+b.w;}));
            var top=Math.min.apply(null,out.map(function(b){return b.y;})),bottom=Math.max.apply(null,out.map(function(b){return b.y+b.h;}));
            if(mode==='horizontal'||mode==='vertical') {
                if(out.length<3)fail('Distribution needs at least three shapes.');
                var horizontal=mode==='horizontal',position=horizontal?'x':'y',size=horizontal?'w':'h',start=horizontal?left:top,end=horizontal?right:bottom;
                out.sort(function(a,b){return a[position]-b[position];});
                var total=out.reduce(function(n,b){return n+b[size];},0),gap=(end-start-total)/(out.length-1);
                if(gap<0)fail('There is not enough space to distribute without overlap.');
                out.forEach(function(b){b[position]=start;start+=b[size]+gap;});
            } else out.forEach(function(b){
                if(mode==='left')b.x=left;else if(mode==='center')b.x=(left+right-b.w)/2;else if(mode==='right')b.x=right-b.w;
                else if(mode==='top')b.y=top;else if(mode==='middle')b.y=(top+bottom-b.h)/2;else if(mode==='bottom')b.y=bottom-b.h;
            });
            return out;
        }
        function signature(raw) {
            // extX/extY on serialized shapes are derived cached values. Canonical
            // spPr.xfrm extents/offsets and every text/style/notes field remain exact.
            return JSON.stringify(JSON.parse(raw),function(key,value){
                if((key==='extX'||key==='extY')&&this&&this.type==='shape'&&this.spPr)return undefined;
                if(key==='rgba'&&this&&['uniColor','scheme','sys'].indexOf(this.type)>=0)return undefined;
                return value;
            });
        }
        function themeSignature(raw){
            return JSON.stringify(JSON.parse(raw),function(key,value){if(key==='rgba'&&this&&['uniColor','scheme','sys'].indexOf(this.type)>=0)return undefined;return value;});
        }
        return {names:names,mutating:mutating,validate:validate,paragraphs:paragraphs,notes:notes,align:align,signature:signature,themeSignature:themeSignature};
    };
});

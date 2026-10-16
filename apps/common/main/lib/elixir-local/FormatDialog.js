/*
 * SPDX-FileCopyrightText: 2026 ElixirOffice contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 */
(function(root,factory){
 'use strict';
 var api=factory();
 root.EOFormatDialog=api;
 if(typeof define==='function'&&define.amd)define([],function(){return api;});
 if(typeof module==='object'&&module.exports)module.exports=api;
})(typeof window==='undefined'?globalThis:window,function(){
 'use strict';
 var sequence=0;
 var css=[
 '.eo-format-dialog{box-sizing:border-box;width:520px;max-width:calc(100vw - 32px);max-height:calc(100vh - 32px);margin:auto;padding:0;border:1px solid var(--eo-dialog-border,#c9c9c9);border-radius:8px;background:var(--eo-dialog-bg,#fff);color:var(--eo-dialog-text,#242424);box-shadow:0 12px 40px #0003;font:14px/1.4 "Segoe UI",Arial,sans-serif;text-align:start;color-scheme:light}',
 '.eo-format-dialog::backdrop{background:rgba(0,0,0,.38)}',
 '.eo-format-dialog[data-theme=dark]{--eo-dialog-bg:#292929;--eo-dialog-text:#f5f5f5;--eo-dialog-muted:#d0d0d0;--eo-dialog-border:#626262;--eo-dialog-selected:#303d45;--eo-dialog-hover:#383838;color-scheme:dark}',
 '.eo-format-dialog *{box-sizing:border-box}',
 '.eo-format-dialog header{padding:20px 24px 12px}',
 '.eo-format-dialog h2{margin:0;font-size:20px;line-height:1.3;font-weight:600}',
 '.eo-format-file{padding:0 24px 12px;display:block}',
 '.eo-format-file label{display:block;font-weight:600;margin-bottom:6px}',
 '.eo-format-file input{width:100%;min-height:34px;padding:6px 10px;border:1px solid var(--eo-dialog-border,#bdbdbd);border-radius:4px;background:var(--eo-dialog-bg,#fff);color:inherit;font:inherit}',
 '.eo-format-file input:focus-visible{outline:2px solid var(--eo-dialog-accent,#0078d4);outline-offset:1px}',
 '.eo-format-error{color:var(--eo-dialog-text,#242424);font-size:13px;margin:6px 0 0}',
 '.eo-format-dialog fieldset{min-width:0;margin:0;padding:0 24px;border:0}',
 '.eo-format-dialog legend{width:100%;padding:0 0 10px;border:0;color:var(--eo-dialog-text,#242424);font-size:14px;font-weight:600}',
 '.eo-format-list{max-height:min(360px,52vh);overflow:auto;padding:3px;border:1px solid var(--eo-dialog-border,#c9c9c9);border-radius:4px}',
 '.eo-format-row{display:flex;align-items:flex-start;gap:12px;min-height:62px;padding:10px 12px;border:1px solid transparent;border-radius:3px;cursor:pointer;text-align:start}',
 '.eo-format-row:hover{background:var(--eo-dialog-hover,#f6f6f6)}',
 '.eo-format-row[data-selected=true]{background:var(--eo-dialog-selected,#eaf3fb);border-color:var(--eo-dialog-accent,#0078d4)}',
 '.eo-format-row:focus-within{outline:2px solid var(--eo-dialog-accent,#0078d4);outline-offset:-2px}',
 '.eo-format-row input{flex:none;margin:4px 0 0;width:16px;height:16px;accent-color:var(--eo-dialog-accent,#0078d4)}',
 '.eo-format-copy{display:block;min-width:0;overflow-wrap:anywhere}',
 '.eo-format-name{display:block;font-weight:600}',
 '.eo-format-description{display:block;margin-top:2px;color:var(--eo-dialog-muted,#616161);font-size:13px}',
 '.eo-format-dialog footer{display:flex;justify-content:flex-end;gap:8px;padding:20px 24px}',
 '.eo-format-dialog button{min-width:84px;min-height:34px;padding:6px 18px;border:1px solid var(--eo-dialog-border,#bdbdbd);border-radius:4px;background:var(--eo-dialog-bg,#fff);color:inherit;font:inherit;font-weight:600;cursor:pointer}',
 '.eo-format-dialog button:hover{background:var(--eo-dialog-hover,#f2f2f2)}',
 '.eo-format-dialog button.eo-format-primary{border-color:var(--eo-dialog-accent,#0078d4);background:var(--eo-dialog-accent,#0078d4);color:#fff}',
 '.eo-format-dialog button:focus-visible{outline:2px solid var(--eo-dialog-text,#242424);outline-offset:3px}',
 '@media(max-width:420px){.eo-format-dialog header{padding:16px 16px 12px}.eo-format-file{padding:0 16px 12px}.eo-format-dialog fieldset{padding:0 16px}.eo-format-dialog footer{padding:16px}.eo-format-row{padding:10px 8px}}',
 '@media(forced-colors:active){.eo-format-row[data-selected=true]{border:2px solid Highlight}.eo-format-dialog button.eo-format-primary{background:Highlight;color:HighlightText}}'
 ].join('\n');
 function open(options,fileChoice){
  var doc=options.document||document;
  if(!Array.isArray(options.formats)||!options.formats.length||options.formats.length>20)throw Error('No available Save As formats');
  if(doc.querySelector('.eo-format-dialog'))return Promise.resolve(null);
  var values=options.formats.map(function(item){return item.value;});
  if(values.some(function(value){return typeof value!=='string'||!/^[a-z0-9-]{1,20}$/.test(value);})||new Set(values).size!==values.length)throw Error('Invalid Save As formats');
  if(!doc.getElementById('eo-format-dialog-style')){
   var style=doc.createElement('style');style.id='eo-format-dialog-style';style.textContent=css;doc.head.appendChild(style);
  }
  return new Promise(function(resolve,reject){
   var previous=doc.activeElement,box=doc.createElement('dialog'),form=doc.createElement('form'),header=doc.createElement('header'),title=doc.createElement('h2');
   var fieldset=doc.createElement('fieldset'),legend=doc.createElement('legend'),list=doc.createElement('div'),footer=doc.createElement('footer');
   var group='eo-format-'+(++sequence),selected=values.indexOf(options.defaultFormat)>=0?options.defaultFormat:values[0],finished=false;
   if(typeof options.accent==='string'&&/^(#[a-f0-9]{3,8}|rgba?\([0-9.,% ]+\))$/i.test(options.accent))box.style.setProperty('--eo-dialog-accent',options.accent);
   box.className='eo-format-dialog';box.setAttribute('aria-modal','true');
   title.id=group+'-title';title.textContent=options.title;box.setAttribute('aria-labelledby',title.id);
   box.lang=options.language||doc.documentElement.lang||'en';
   box.dir=options.rtl||doc.documentElement.dir==='rtl'||doc.body.classList.contains('rtl')||/^(ar|he|ur|fa)(-|$)/i.test(box.lang)?'rtl':'ltr';
   var dark=options.dark;
   if(dark===undefined)dark=/dark/.test(doc.body.className+' '+doc.documentElement.className)||doc.defaultView.matchMedia('(prefers-color-scheme: dark)').matches;
   box.dataset.theme=dark?'dark':'light';
   header.appendChild(title);form.appendChild(header);
   var fileInput=null,fileError=null,fileOptions=options.fileName;
   if(fileChoice&&fileOptions){
    var fileBox=doc.createElement('div'),fileLabel=doc.createElement('label');
    fileBox.className='eo-format-file';fileInput=doc.createElement('input');fileInput.type='text';
    fileInput.id=group+'-name';fileInput.value=fileOptions.value||'';fileInput.maxLength=fileOptions.maxLength||200;
    fileInput.autocomplete='off';fileInput.spellcheck=false;
    fileLabel.htmlFor=fileInput.id;fileLabel.textContent=fileOptions.label;
    fileError=doc.createElement('p');fileError.className='eo-format-error';fileError.id=group+'-name-error';fileError.setAttribute('role','alert');fileError.hidden=true;
    fileInput.setAttribute('aria-describedby',fileError.id);
    fileInput.addEventListener('input',function(){fileError.hidden=true;fileInput.removeAttribute('aria-invalid');});
    fileBox.appendChild(fileLabel);fileBox.appendChild(fileInput);fileBox.appendChild(fileError);form.appendChild(fileBox);
   }
   legend.textContent=options.label;fieldset.appendChild(legend);
   list.className='eo-format-list';fieldset.appendChild(list);form.appendChild(fieldset);
   var radios=[];
   options.formats.forEach(function(item){
    var row=doc.createElement('label'),radio=doc.createElement('input'),copy=doc.createElement('span'),name=doc.createElement('span'),description=doc.createElement('span');
    row.className='eo-format-row';row.dataset.selected=String(item.value===selected);
    radio.type='radio';radio.name=group;radio.value=item.value;radio.checked=item.value===selected;
    copy.className='eo-format-copy';name.className='eo-format-name';description.className='eo-format-description';
    var suffix=String(item.name).match(/^(.*?)(\(\*\.[a-z0-9]+\))$/i);
    if(suffix){name.appendChild(doc.createTextNode(suffix[1]));var extension=doc.createElement('bdi');extension.dir='ltr';extension.textContent=suffix[2];name.appendChild(extension);}
    else name.textContent=item.name;
    name.id=group+'-'+item.value+'-name';description.textContent=item.description;description.id=group+'-'+item.value+'-description';
    radio.setAttribute('aria-labelledby',name.id);
    radio.setAttribute('aria-describedby',description.id);
    copy.appendChild(name);copy.appendChild(description);row.appendChild(radio);row.appendChild(copy);list.appendChild(row);radios.push(radio);
    radio.addEventListener('change',function(){
     if(!radio.checked)return;selected=item.value;
     Array.prototype.forEach.call(list.children,function(child){child.dataset.selected=String(child===row);});
     row.scrollIntoView({block:'nearest'});
    });
   });
   var ok=doc.createElement('button'),cancel=doc.createElement('button');
   ok.type='submit';ok.className='eo-format-primary';ok.textContent=options.ok;
   cancel.type='button';cancel.textContent=options.cancel;
   if(options.platform==='darwin'){footer.appendChild(cancel);footer.appendChild(ok);}else{footer.appendChild(ok);footer.appendChild(cancel);}
   form.appendChild(footer);box.appendChild(form);
   function finish(value){
    if(finished)return;
    var name=fileInput&&fileInput.value;
    if(value!==null&&fileInput&&(!name.trim()||fileOptions.validate&&!fileOptions.validate(name,value))){
     fileError.textContent=fileOptions.error;fileError.hidden=false;fileInput.setAttribute('aria-invalid','true');fileInput.focus();return;
    }
    finished=true;
    if(box.open)box.close();box.remove();
    if(options.signal)options.signal.removeEventListener('abort',onAbort);
    if(previous&&previous.isConnected&&typeof previous.focus==='function')previous.focus();
    resolve(value!==null&&fileChoice?{format:value,name:name}:value);
   }
   cancel.addEventListener('click',function(){finish(null);});
   form.addEventListener('submit',function(event){event.preventDefault();finish(selected);});
   box.addEventListener('cancel',function(event){event.preventDefault();finish(null);});
   box.addEventListener('close',function(){finish(null);});
   box.addEventListener('keydown',function(event){
    // Editor/File-menu shortcuts must not consume modal keys.
    event.stopPropagation();
    if(event.isComposing||event.keyCode===229)return;
    if(event.key==='Escape'){event.preventDefault();finish(null);return;}
    if(event.key==='Tab'){
     var controls=(fileInput?[fileInput]:[]).concat([box.querySelector('input:checked')]).concat(Array.prototype.slice.call(footer.children));
     var first=controls[0],last=controls[controls.length-1];
     if(event.shiftKey&&event.target===first){event.preventDefault();last.focus();}
     else if(!event.shiftKey&&event.target===last){event.preventDefault();first.focus();}
    }
    if(event.key==='Enter'){event.preventDefault();finish(event.target===cancel?null:selected);}
   });
   function onAbort(){finish(null);}
   if(options.signal){
    if(options.signal.aborted){finish(null);return;}
    options.signal.addEventListener('abort',onAbort,{once:true});
   }
   doc.body.appendChild(box);
   try{box.showModal();radios[values.indexOf(selected)].focus();}
   catch(error){finished=true;box.remove();if(options.signal)options.signal.removeEventListener('abort',onAbort);if(previous&&previous.isConnected)previous.focus();reject(error);}
  });
 }
 return {open:open,chooseFile:function(options){return open(options,true);}};
});

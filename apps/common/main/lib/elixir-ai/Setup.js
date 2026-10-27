// SPDX-License-Identifier: AGPL-3.0-only
define(['common/main/lib/elixir-ai/Core','common/main/lib/elixir-ai/View','common/main/lib/elixir-ai/SetupStrings','common/main/lib/elixir-ai/ProviderCatalog'],function (ElixirAI,_view,English,Providers) {
 'use strict';
 var catalog=Object.freeze([]);
 function https(value){return typeof value==='string'&&value.length<=2048&&/^https:\/\//i.test(value)&&ElixirAI.Settings.endpointAllowed(value,false);}
 ElixirAI.SetupCatalog=Object.freeze({
  configure:function(values){
   if(!Array.isArray(values)||values.length>32)throw Error('Invalid AI provider catalog');
   var next=values.map(function(value){
    if(!value||typeof value.label!=='string'||!value.label.trim()||value.label.length>80||!https(value.endpoint)||value.model!==undefined&&(typeof value.model!=='string'||value.model.length>200))throw Error('Invalid AI provider catalog');
    return Object.freeze({label:value.label,endpoint:value.endpoint,model:value.model||''});
   });
   catalog=Object.freeze(next);
  },
  get:function(){return catalog;}
 });
 ElixirAI.SetupCatalog.configure(Providers||[]);
 ElixirAI.SetupStrings=ElixirAI.SetupStrings||{};
 function text(key){var localized=ElixirAI.SetupStrings||{};return typeof localized[key]==='string'&&localized[key].trim()?localized[key]:English[key];}
 function SetupAssistant(host){this.host=host||{};this.current=null;this.disposed=false;}
 SetupAssistant.prototype.open=function(invoker){
  var me=this,doc=document;
  if(me.disposed)return false;
  if(me.current){me.current.endpoint.focus();return true;}
  function node(tag,parent,value){
   var el=doc.createElement(tag);if(value!==undefined)el.textContent=value;
   el.setAttribute('oo_editor_input','true');if(parent)parent.appendChild(el);return el;
  }
  var dialog=node('dialog');dialog.id='elixir-ai-setup-dialog';dialog.setAttribute('aria-labelledby','elixir-ai-setup-title');dialog.setAttribute('aria-modal','true');
  dialog.className='elixir-ai-dlg';dialog.style.cssText='box-sizing:border-box;width:min(600px,calc(100vw - 32px));max-height:calc(100vh - 32px);overflow:auto;padding:24px;border:1px solid var(--border-regular-control,#d1d1d1);border-radius:8px;background:var(--background-normal,#fff);color:var(--text-normal,#242424);font:inherit';
  var title=node('h2',dialog,text('title'));title.id='elixir-ai-setup-title';
  node('p',dialog,text('intro'));node('p',dialog,text('disclosure'));node('p',dialog,text('noChecks'));
  function field(id,label,type){
   var wrapper=node('div',dialog),caption=node('label',wrapper,label);caption.setAttribute('for',id);
   var input=node(type==='select'?'select':'input',wrapper);input.id=id;
   if(type!=='select')input.setAttribute('type',type);input.value='';input.style.cssText='display:block;box-sizing:border-box;width:100%;margin:4px 0 12px;padding:6px';return input;
  }
  var provider=field('elixir-ai-setup-provider',text('provider'),'select'),generic=node('option',provider,text('generic'));generic.value='custom';provider.value='custom';
  var providers=ElixirAI.SetupCatalog.get();
  providers.forEach(function(value,index){node('option',provider,value.label).value=String(index);});
  var endpoint=field('elixir-ai-setup-endpoint',text('endpoint'),'url'),model=field('elixir-ai-setup-model',text('model'),'text'),key=field('elixir-ai-setup-key',text('key'),'password');key.setAttribute('autocomplete','off');
  node('p',dialog,text(window.eoDesktopAI?'keyDesktop':'keyBrowser'));node('p',dialog,text('destination'));
  var error=node('p',dialog);error.id='elixir-ai-setup-error';error.setAttribute('role','status');
  var actions=node('div',dialog);actions.className='elixir-ai-actions';
  var cancel=node('button',actions,text('cancel'));cancel.id='elixir-ai-setup-cancel';cancel.setAttribute('type','button');cancel.className='elixir-ai-btn';
  var proceed=node('button',actions,text('continue'));proceed.id='elixir-ai-setup-continue';proceed.setAttribute('type','button');proceed.className='elixir-ai-btn primary';
  var state={dialog:dialog,endpoint:endpoint,key:key,invoker:invoker||doc.activeElement,busy:false};me.current=state;
  provider.addEventListener('change',function(){
   var index=provider.value,value=/^\d+$/.test(index)?providers[Number(index)]:null;
   endpoint.value=value?value.endpoint:'';model.value=value?value.model:'';error.textContent='';endpoint.focus();
  });
  cancel.addEventListener('click',function(){me.close(state);});
  proceed.addEventListener('click',async function(){
   if(me.disposed||me.current!==state||state.busy)return;
   var url=endpoint.value.trim(),name=model.value.trim();
   if(!https(url)||!name||name.length>200||key.value.length>8192||/[\r\n]/.test(key.value)){error.textContent=text('invalid');return;}
   state.busy=true;proceed.disabled=true;error.textContent=text('waiting');
   var value={endpoint:url,model:name,apiKey:key.value};key.value='';
   try{
    var saved=await ElixirAI.Settings.save(value);value.apiKey='';
    if(me.disposed||me.current!==state)return;
    if(saved===false){error.textContent=text('denied');return;}
    me.close(state);if(me.host.onSettingsChanged)me.host.onSettingsChanged();
   }catch(_){value.apiKey='';if(me.current===state&&!me.disposed)error.textContent=text('failed');}
   finally{state.busy=false;proceed.disabled=false;}
  });
  dialog.addEventListener('cancel',function(event){event.preventDefault();me.close(state);});
  dialog.addEventListener('close',function(){me.close(state);});
  dialog.addEventListener('keydown',function(event){event.stopPropagation();if(event.key==='Escape'){event.preventDefault();me.close(state);}});
  ['keyup','keypress'].forEach(function(name){dialog.addEventListener(name,function(event){event.stopPropagation();});});
  doc.body.appendChild(dialog);dialog.showModal();endpoint.focus();return true;
 };
 SetupAssistant.prototype.close=function(state){
  state=state||this.current;if(!state||this.current!==state)return;
  this.current=null;state.key.value='';if(state.dialog.open)state.dialog.close();state.dialog.remove();
  var target=state.invoker;
  if(!this.disposed&&target&&target.isConnected&&(!target.getClientRects||target.getClientRects().length))target.focus();
  else if(!this.disposed){var input=document.getElementById('area_id');if(input)input.focus();}
 };
 SetupAssistant.prototype.dispose=function(){this.disposed=true;this.close();};
 ElixirAI.SetupAssistant=SetupAssistant;
 return ElixirAI;
});

/*
 * SPDX-FileCopyrightText: 2026 ElixirOffice contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 */
(function(factory){
 if(typeof define==='function'&&define.amd)define([],factory);else factory();
})(function(){
 var window=globalThis;
 'use strict';
 function guardNetwork(){
  function allowed(url,method){
   var u=new URL(url,window.location.href);
   return (u.protocol==='blob:'||u.protocol==='data:') ||
    ((method||'GET').toUpperCase()==='GET'&&u.origin===window.location.origin&&/^\/example\/local-assets\/(?:\d+\.\d+\.\d+-[a-f0-9]+\/)?(sdkjs|web-apps|fonts|dictionaries)\//.test(u.pathname));
  }
  function blocked(){return new Error('This command requires a server and is unavailable in local mode');}
  var fetch=window.fetch,open=XMLHttpRequest.prototype.open;
  window.fetch=function(input,options){
   var url=typeof input==='string'||input instanceof URL?input:input.url;
   var method=options&&options.method||input.method||'GET';
   return allowed(url,method)?fetch.call(window,input,options):Promise.reject(blocked());
  };
  XMLHttpRequest.prototype.open=function(method,url){if(!allowed(url,method))throw blocked();return open.apply(this,arguments);};
  window.WebSocket=function(){throw blocked();};
  window.EventSource=function(){throw blocked();};
  if(window.navigator.sendBeacon)window.navigator.sendBeacon=function(){return false;};
 }
 function prepare(api,main){
  guardNetwork();
  api._openChartOrLocalDocument=function(){};
  api.asc_isOffline=function(){return true;};
  var permissions=main.onEditorPermissions;
  main.onEditorPermissions=function(params){
   if(api.DocInfo&&api.DocInfo.get_OfflineApp()){
    // Local editing uses this application's bundled edition, without a remote
    // license/version handshake or collaboration capabilities.
    params.setLicenseType(Asc.c_oLicenseResult.Success);
    params.setRights(Asc.c_oRights.Edit);params.setIsLight(true);
    var version=main.getApplication().getController('LeftMenu').leftMenu.getMenu('about').txtVersionNum.match(/^(\d+\.\d+\.\d+)/);
    params.setBuildVersion(version&&version[1]);
   }
   return permissions.call(this,params);
  };
 }
 function install(api,options){
  options=options||{};var resources=[],urls=[],saving=null,pendingImage=Promise.resolve(),format=options.format;
  var editorFormat={docx:'doct',odt:'doct',xlsx:'xlst',ods:'xlst',pptx:'pptt',odp:'pptt'}[format];
  if(!editorFormat)throw new Error('Unsupported local document format');
  if(!api.DocInfo||!api.DocInfo.get_OfflineApp())throw new Error('Local mode requires an offline editor');
  function signal(){return options.getSignal?options.getSignal():options.signal;}
  function serialize(){
   if(editorFormat==='doct')return new AscCommonWord.BinaryFileWriter(api.WordControl.m_oLogicDocument).Write(true);
   if(editorFormat==='xlst'){api.wbModel.dependencyFormulas.calcTree();return new AscCommonExcel.BinaryFileWriter(api.wbModel).Write(true);}
   api.WordControl.m_oLogicDocument.CalculateComments();return new AscCommon.CBinaryFileWriter().WriteDocument3(api.WordControl.m_oLogicDocument,false);
  }
  async function save(){
   if(saving)return saving;
   saving=(async function(){
    if(options.onBeforeSave)options.onBeforeSave();
    await pendingImage;
    var bytes=serialize();
    var output=await window.EOWasmConvert.convert(bytes,editorFormat,format,{resources:resources,onProgress:options.onProgress,signal:signal()});
    if(options.onSave)await options.onSave(output);
    return output;
   })().finally(function(){saving=null;if(options.onAfterSave)options.onAfterSave();});
   return saving;
  }
  function report(e){if(options.onError)options.onError(e);}
  // Only this explicitly local API instance is patched; hosted editors keep their methods.
  api.asc_Save=function(){save().catch(report);return true;};
  api.downloadAs=function(action,downloadOptions){
   if(downloadOptions&&downloadOptions.fileType!==window.EOWasmConvert.formats[format]){
    report(new Error('This download format is unavailable in local mode. Save as .'+format+'.'));return;
   }
   save().catch(report);
  };
  api.asc_isOffline=function(){return true;};
  api.isAutoSave=false;
  var docUrls=AscCommon.g_oDocumentUrls,add=docUrls.addImageUrl;
  var imageUrls=new Map();
  docUrls.addImageUrl=function(name,url){return add.call(this,name,imageUrls.get(name)||url);};
  async function addImage(file,imageOptions){
   var extension={ 'image/png':'png','image/jpeg':'jpg' }[file.type];
   if(!extension)throw new Error('Local pictures must be PNG or JPEG');
   var bytes=new Uint8Array(await file.arrayBuffer()),name='local-'+crypto.randomUUID()+'.'+extension;
   var url=URL.createObjectURL(new Blob([bytes],{type:file.type}));urls.push(url);
   resources.push({path:'media/'+name,bytes:bytes});imageUrls.set(name,url);
   add.call(docUrls,name,url);
   await new Promise(function(resolve){api.ImageLoader.LoadImagesWithCallback([url],resolve,[]);});
   api._addImageUrl([url],imageOptions);
  }
  api.asc_addImage=function(imageOptions){
   var picker=document.createElement('input');picker.type='file';picker.accept='image/png,image/jpeg';
   picker.addEventListener('change',function(){if(picker.files[0])pendingImage=addImage(picker.files[0],imageOptions).catch(report);},{once:true});
   picker.click();
  };
  async function open(input){
   var output=await window.EOWasmConvert.convert(new Uint8Array(input),format,editorFormat,{onProgress:options.onProgress,signal:signal()});
   resources=output.resources||[];resources.forEach(function(f){
    var url=URL.createObjectURL(new Blob([f.bytes]));urls.push(url);
    imageUrls.set(f.path.replace(/^media\//,''),url);imageUrls.set(f.path,url);
    add.call(docUrls,f.path.replace(/^media\//,''),url);
   });
   await new Promise(function(resolve,reject){
    var currentSignal=signal();
    var timer=setTimeout(function(){finish(new Error('Local document did not become ready'));},90000);
    function finish(e){clearTimeout(timer);api.asc_unregisterCallback('asc_onDocumentContentReady',ready);if(currentSignal)currentSignal.removeEventListener('abort',cancel);e?reject(e):resolve();}
    function ready(){finish();}
    function cancel(){finish(new Error('Opening cancelled'));}
    api.asc_registerCallback('asc_onDocumentContentReady',ready);
    if(currentSignal)currentSignal.addEventListener('abort',cancel,{once:true});
    try{api.asc_openDocumentFromBytes(output);}catch(e){finish(e);}
   });
   return output.stats;
  }
  return {open:open,save:save,serialize:serialize,addImage:addImage,dispose:function(){urls.forEach(URL.revokeObjectURL);docUrls.addImageUrl=add;}};
 }
 window.Common=window.Common||{};window.Common.ElixirLocal={prepare:prepare,install:install};
 return window.Common.ElixirLocal;
});

/*
 * SPDX-FileCopyrightText: 2026 ElixirOffice contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 */
(function(factory){
 if(typeof define==='function'&&define.amd)define(['common/main/lib/elixir-local/Images','common/main/lib/elixir-local/Pdf','common/main/lib/elixir-local/Password'],factory);else factory(window.Common&&window.Common.ElixirLocalImages,window.Common&&window.Common.ElixirLocalPdf,window.Common&&window.Common.ElixirLocalPassword);
})(function(Images,Pdf,Password){
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
 function prepare(api,main,capabilities){
  guardNetwork();
  api.eoLocalMode=true;
  Password.prepare(api,main);
  api._openChartOrLocalDocument=function(){};
  api.asc_isOffline=function(){return true;};
  var permissions=main.onEditorPermissions;
  main.onEditorPermissions=function(params){
   if(api.DocInfo&&api.DocInfo.get_OfflineApp()){
    // Local editing uses this application's bundled edition, without a remote
    // license/version handshake or collaboration capabilities.
    params.setLicenseType(Asc.c_oLicenseResult.Success);
    params.setRights(Asc.c_oRights.Edit);params.setIsLight(true);
    // Single-user review is independent of the collaboration edition.
    main.appOptions.canLocalComments=true;main.appOptions.canLocalReview=true;
    main.appOptions.canFeatureComparison=false;
    main.appOptions.isElixirLocal=true;
    if(capabilities)main.appOptions.localDownloadFormats=capabilities.downloadFormats;
    var version=main.getApplication().getController('LeftMenu').leftMenu.getMenu('about').txtVersionNum.match(/^(\d+\.\d+\.\d+)/);
    params.setBuildVersion(version&&version[1]);
   }
   return permissions.call(this,params);
  };
 }
 function install(api,options){
  options=options||{};if(options.format==='pdf')return Pdf.install(api,options);var resources=[],urls=[],saving=null,pendingImage=Promise.resolve(),pendingImages=0,disposed=false,revision=0,format=options.format,exporting=null;
  var editorFormat={docx:'doct',odt:'doct',xlsx:'xlst',ods:'xlst',pptx:'pptt',odp:'pptt'}[format];
  if(!editorFormat)throw new Error('Unsupported local document format');
  if(!api.DocInfo||!api.DocInfo.get_OfflineApp())throw new Error('Local mode requires an offline editor');
  Images.assertAvailable();
  var passwords=Password.install(api,Object.assign({},options,{save:save,isBusy:function(){return !!(saving||exporting);},onError:function(e){if(options.onError)options.onError(e);}}));
  var wordHistoryState;
  function action(){revision++;}
  function modified(){
   if(!api.isDocumentModified())return;
   if(editorFormat!=='doct'){revision++;return;}
   // Word repeats dirty notifications during layout without a document edit.
   var history=AscCommon.History,point=history.Points[history.Index];
   var state=history.Index+':'+history.Points.length+':'+(point?point.Items.length:0);
   if(state!==wordHistoryState){wordHistoryState=state;revision++;}
  }
  api.asc_registerCallback('asc_onDocumentModifiedChanged',modified);
  var actions=editorFormat==='doct'?['asc_onUserActionStart','asc_onBeforeUndoRedo']:[];
  actions.forEach(function(event){api.asc_registerCallback(event,action);});
  function signal(){return options.getSignal?options.getSignal():options.signal;}
  function serialize(){
   if(editorFormat==='doct')return new AscCommonWord.BinaryFileWriter(api.WordControl.m_oLogicDocument).Write(true);
   if(editorFormat==='xlst'){api.wbModel.dependencyFormulas.calcTree();return new AscCommonExcel.BinaryFileWriter(api.wbModel).Write(true);}
   api.WordControl.m_oLogicDocument.CalculateComments();return new AscCommon.CBinaryFileWriter().WriteDocument3(api.WordControl.m_oLogicDocument,false);
  }
  async function save(passwordChoice){
   if(saving){if(passwordChoice!==undefined)throw new Error('A local save is already running');return saving;}
   if(exporting)throw new Error('A local export is already running');
   saving=(async function(){
    if(options.onBeforeSave)options.onBeforeSave();
    await pendingImage;
    var saveRevision=revision,bytes=serialize();
    var output=await passwords.save(async function(passwordOptions,protection){
     var output=await window.EOWasmConvert.convert(bytes,editorFormat,format,Object.assign({resources:resources,onProgress:options.onProgress,signal:signal()},passwordOptions));
     if(options.onSave)await options.onSave(output,protection);return output;
    },passwordChoice);
    // Mark only the completed snapshot saved; edits made during conversion stay dirty.
    if(revision===saveRevision){AscCommon.History.Reset_SavedIndex(true);api.SetDocumentModified(false);}
    else if(editorFormat==='doct'){
     // A concurrent undo can return to the old saved state after this snapshot
     // overwrites it. Keep that unsaved state dirty through Word's layout checks.
     AscCommon.History.ForceSave=true;api.CheckChangedDocument();
    }
    return output;
   })().finally(function(){saving=null;if(options.onAfterSave)options.onAfterSave();});
   return saving;
  }
  function report(e){if(options.onError)options.onError(e);}
  // Only this explicitly local API instance is patched; hosted editors keep their methods.
  // Automatic server saves must never overwrite a local user's file.
  // Desktop recovery takes its own snapshots; explicit Save retains the normal path.
  api.asc_Save=function(isAutoSave){if(isAutoSave===true)return false;save().catch(report);return true;};
  var outputs={doct:['docx','odt','rtf','txt','pdf'],xlst:['xlsx','ods','csv','pdf'],pptt:['pptx','odp','pdf']}[editorFormat];
  async function exportAs(target,downloadOptions){
   if(outputs.indexOf(target)<0)throw new Error('This download format is unavailable in local mode');
   if(!options.onExport)throw new Error('Use the desktop Save As command to export this format');
   if(saving||exporting)throw new Error('A local conversion is already running');
   exporting=(async function(){
    if(options.onBeforeSave)options.onBeforeSave();
    var state=await snapshot(),output;
    if(target==='pdf'){
     if(!options.buildPdf)throw new Error('The bundled PDF exporter is unavailable');
     await passwords.confirmExport();output=await options.buildPdf(api,format,downloadOptions||{},signal());
    }else{
     var conversion={resources:state.resources,onProgress:options.onProgress,signal:signal()},advanced=downloadOptions&&downloadOptions.advancedOptions;
     if(target==='csv'&&advanced&&advanced.asc_getCodePage){
      var delimiter=advanced.asc_getDelimiter(),custom=advanced.asc_getDelimiterChar();
      if(Array.isArray(delimiter)){
       if(delimiter.length>1)throw new Error('CSV export requires one delimiter');
       delimiter=delimiter[0];
      }
      if(delimiter===-1&&custom)delimiter=0;
      conversion.csv={encoding:advanced.asc_getCodePage(),delimiter:delimiter,delimiterChar:custom};
     }
     output={bytes:await passwords.copyAs(target,function(passwordOptions){return window.EOWasmConvert.convert(state.bytes,state.format,target,Object.assign(conversion,passwordOptions));})};
    }
    if(disposed||(signal()&&signal().aborted))throw new Error('Conversion cancelled');
    await options.onExport(output.bytes,target,downloadOptions||{},output.stats);
    // A download is a copy: keep the editing format, original and undo/saved state.
    return output.bytes;
   })().finally(function(){exporting=null;if(options.onAfterSave)options.onAfterSave();});
   return exporting;
  }
  api.downloadAs=function(action,downloadOptions){
   if(!options.onExport){
    if(downloadOptions&&downloadOptions.fileType!==window.EOWasmConvert.formats[format]){
     report(new Error('Use the desktop Save As command to export this format'));return;
    }
    save().catch(report);return;
   }
   downloadOptions=downloadOptions||{};
   var target=downloadOptions.fileType===Asc.c_oAscFileType.PDF?'pdf':Object.keys(window.EOWasmConvert.formats).find(function(key){return window.EOWasmConvert.formats[key]===downloadOptions.fileType;});
   if(downloadOptions.fileType===undefined)target=format;
   if(outputs.indexOf(target)<0){report(new Error('This download format is unavailable in local mode'));return;}
   if(action)api.sync_StartAction(Asc.c_oAscAsyncActionType.BlockInteraction,action);
   exportAs(target,downloadOptions).catch(report).finally(function(){
    if(action)api.sync_EndAction(Asc.c_oAscAsyncActionType.BlockInteraction,action);
   });
  };
  api.asc_isOffline=function(){return true;};
  api.isAutoSave=false;
  var docUrls=AscCommon.g_oDocumentUrls,add=docUrls.addImageUrl;
  var imageUrls=new Map();
  var localImageUrl=function(name,url){return add.call(this,name,imageUrls.get(name)||url);};
  docUrls.addImageUrl=localImageUrl;
  function enqueueImage(job){
   pendingImages++;
   var operation=pendingImage.then(job);
   pendingImage=operation.catch(function(){}).finally(function(){pendingImages--;});
   return operation;
  }
  var pictures=Images.install(api,{
   signal:signal,enqueue:enqueueImage,onError:report,
   lookup:function(url){var path=docUrls.getLocal(url);return path?{url:url,path:path}:null;},
   register:function(bytes,extension,url){
    var name='local-'+crypto.randomUUID()+'.'+extension;
    urls.push(url);resources.push({path:'media/'+name,bytes:bytes});imageUrls.set(name,url);imageUrls.set('media/'+name,url);add.call(docUrls,name,url);
    return {url:url,path:'media/'+name};
   }
  });
  function addImage(file,imageOptions){return pictures.insert(file,imageOptions);}
  api.asc_addImage=function(imageOptions){
   var picker=document.createElement('input');picker.type='file';picker.accept='image/png,image/jpeg,image/gif,image/bmp,image/svg+xml,image/webp';
   picker.addEventListener('change',function(){if(picker.files[0]){addImage(picker.files[0],imageOptions).catch(report);}},{once:true});
   picker.click();
  };
  async function open(input){
   var source=new Uint8Array(input);
   var output=await passwords.open(source,function(passwordOptions){return window.EOWasmConvert.convert(source,format,editorFormat,Object.assign({onProgress:options.onProgress,signal:signal()},passwordOptions));});
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
    if(currentSignal){currentSignal.addEventListener('abort',cancel,{once:true});if(currentSignal.aborted){cancel();return;}}
    try{api.asc_openDocumentFromBytes(output);}catch(e){finish(e);}
   }).catch(function(e){passwords.dispose();throw e;});
   pictures.activate();
   return output.stats;
  }
  async function exportCopy(target,write){
   if(outputs.indexOf(target)<0||target==='pdf')throw new Error('Unsupported local converted copy');
   if(saving||exporting)throw new Error('A local conversion is already running');
   exporting=(async function(){var state=await snapshot();return passwords.copyAs(target,async function(passwordOptions,protection){
    var output=await window.EOWasmConvert.convert(state.bytes,state.format,target,Object.assign({resources:state.resources,signal:signal()},passwordOptions));
    if(disposed||(signal()&&signal().aborted))throw new Error('Conversion cancelled');
    await write(output,protection);return output;
   });})();
   try{return await exporting;}finally{exporting=null;}
  }
  // Recovery takes a detached snapshot without saving or resetting document history.
  async function snapshot(){
   await pendingImage;
   return {bytes:serialize(),format:editorFormat,resources:resources.map(function(resource){return {path:resource.path,bytes:new Uint8Array(resource.bytes)};})};
  }
  return {open:open,save:save,exportCopy:exportCopy,saveAs:function(){return passwords.saveAs(save);},confirmUnprotectedExport:passwords.confirmExport,isProtected:passwords.isProtected,isRecoveryLocked:passwords.isRecoveryLocked,exportAs:exportAs,snapshot:snapshot,serialize:serialize,addImage:addImage,isModified:function(){return api.isDocumentModified();},hasPendingImages:function(){return pendingImages>0;},dispose:function(){disposed=true;passwords.dispose();pictures.dispose();api.asc_unregisterCallback('asc_onDocumentModifiedChanged',modified);actions.forEach(function(event){api.asc_unregisterCallback(event,action);});urls.forEach(URL.revokeObjectURL);if(docUrls.addImageUrl===localImageUrl)docUrls.addImageUrl=add;}};
 }
 window.Common=window.Common||{};window.Common.ElixirLocal={prepare:prepare,install:install};
 return window.Common.ElixirLocal;
});

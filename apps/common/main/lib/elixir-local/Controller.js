/*
 * SPDX-FileCopyrightText: 2026 ElixirOffice contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 */
(function(factory){
 if(typeof define==='function'&&define.amd)define(['common/main/lib/elixir-local/Images','common/main/lib/elixir-local/Pdf','common/main/lib/elixir-local/Password','common/main/lib/elixir-local/Operation','common/main/lib/elixir-local/FileNames','common/main/lib/elixir-local/MailMerge'],factory);else factory(window.Common&&window.Common.ElixirLocalImages,window.Common&&window.Common.ElixirLocalPdf,window.Common&&window.Common.ElixirLocalPassword,window.Common&&window.Common.ElixirLocalOperation,window.Common&&window.Common.ElixirLocalFileNames,window.Common&&window.Common.ElixirLocalMailMerge);
})(function(Images,Pdf,Password,Operation,FileNames,MailMerge){
 var window=globalThis;
 'use strict';
 var emptyImportedImage='data:image/gif;base64,R0lGODlhAQABAAD/ACwAAAAAAQABAAACADs=';
 function resolveImportedImage(images,name,url){
  return images.get(name)||(/(?:^|\/)_offline_media\//.test(String(url))?emptyImportedImage:url);
 }
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
  api.eoLocalSaveAs=capabilities&&capabilities.saveAs;
  var mailMerge=!!(capabilities&&capabilities.mailMerge);
  if(mailMerge)main.appOptions.canUseMailMerge=true;
  Password.prepare(api,main);
  api._openChartOrLocalDocument=function(){};
  api.asc_isOffline=function(){return true;};
  var permissions=main.onEditorPermissions;
  main.onEditorPermissions=function(params){
   var local=api.DocInfo&&api.DocInfo.get_OfflineApp();
   if(local){
    // Local editing uses this application's bundled edition, without a remote
    // license/version handshake or collaboration capabilities.
    params.setLicenseType(Asc.c_oLicenseResult.Success);
    params.setRights(Asc.c_oRights.Edit);params.setIsLight(true);
    // Single-user review is independent of the collaboration edition.
    main.appOptions.canLocalComments=true;main.appOptions.canLocalReview=true;
    main.appOptions.canFeatureComparison=false;
    main.appOptions.isElixirLocal=true;
    if(capabilities&&capabilities.mailMerge)main.appOptions.canUseMailMerge=true;
    if(capabilities)main.appOptions.localDownloadFormats=capabilities.downloadFormats;
    var version=main.getApplication().getController('LeftMenu').leftMenu.getMenu('about').txtVersionNum.match(/^(\d+\.\d+\.\d+)/);
    params.setBuildVersion(version&&version[1]);
   }
   var result=permissions.call(this,params);
   if(local&&mailMerge){
    main.appOptions.canUseMailMerge=true;
    Common.NotificationCenter.trigger('elixir:mailmerge-ready');
   }
   return result;
  };
  if(mailMerge)Common.NotificationCenter.trigger('elixir:mailmerge-ready');
 }
 function install(api,options){
  options=options||{};if(options.format==='pdf')return Pdf.install(api,options);var resources=[],urls=[],saving=null,pendingImage=Promise.resolve(),pendingImages=0,disposed=false,revision=0,format=options.format,exporting=null,opening=false,activeJob,externalOpeningJob;
  var editorFormat={docx:'doct',odt:'doct',xlsx:'xlst',ods:'xlst',pptx:'pptt',odp:'pptt'}[format];
  if(!editorFormat)throw new Error('Unsupported local document format');
  if(!api.DocInfo||!api.DocInfo.get_OfflineApp())throw new Error('Local mode requires an offline editor');
  Images.assertAvailable();
  var operations=Operation.create(options.onOperation,options.operationScope);
  function openingBusy(){return opening||externalOpeningJob&&externalOpeningJob.isRunning&&externalOpeningJob.isRunning();}
  function begin(kind){activeJob=operations.begin(kind,{name:api.documentTitle||options.name||''});return activeJob;}
  function progress(job,value){job.progress(value);if(options.onProgress)options.onProgress(value);}
  var passwords=Password.install(api,Object.assign({},options,{save:save,isBusy:function(){return !!(saving||exporting||openingBusy());},onState:function(value){if(activeJob)activeJob.progress(value);},onError:function(e){if(options.onError)options.onError(e);}}));
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
  function serialize(logicDocument){
   if(editorFormat==='doct')return new AscCommonWord.BinaryFileWriter(logicDocument||api.WordControl.m_oLogicDocument).Write(true);
   if(editorFormat==='xlst'){api.wbModel.dependencyFormulas.calcTree();return new AscCommonExcel.BinaryFileWriter(api.wbModel).Write(true);}
   api.WordControl.m_oLogicDocument.CalculateComments();return new AscCommon.CBinaryFileWriter().WriteDocument3(api.WordControl.m_oLogicDocument,false);
  }
  async function save(passwordChoice,existingJob,prepared){
   if(openingBusy())throw Object.assign(new Error('A local operation is already running'),{code:'operation-busy'});
   if(saving){if(passwordChoice!==undefined)throw new Error('A local save is already running');return saving;}
   if(exporting)throw new Error('A local export is already running');
   var job=existingJob||begin('save');
   saving=(async function(){
    if(!prepared&&options.onBeforeSave)options.onBeforeSave();
    if(!prepared&&options.beforeOperationWork)await options.beforeOperationWork();
    if(disposed||(signal()&&signal().aborted))throw new Error('Conversion cancelled');
    await pendingImage;
    var saveRevision=revision,bytes=serialize();
    var output=await passwords.save(async function(passwordOptions,protection){
     var output=await window.EOWasmConvert.convert(bytes,editorFormat,format,Object.assign({resources:resources,onProgress:function(value){progress(job,Object.assign({},value,passwordOptions.savePassword!==undefined&&value.stage==='converting'?{purpose:'encrypting'}:{}));},signal:signal()},passwordOptions));
     job.phase('awaiting-output');
     var acknowledgement=options.onSave?await options.onSave(output,protection):undefined;
     if(acknowledgement&&acknowledgement.kind==='cancelled'){var cancelled=new Error('Save cancelled');cancelled.code='cancelled';throw cancelled;}
     job.finish(acknowledgement);return output;
    },passwordChoice);
    // Mark only the completed snapshot saved; edits made during conversion stay dirty.
    if(revision===saveRevision){AscCommon.History.Reset_SavedIndex(true);api.SetDocumentModified(false);}
    else if(editorFormat==='doct'){
     // A concurrent undo can return to the old saved state after this snapshot
     // overwrites it. Keep that unsaved state dirty through Word's layout checks.
     AscCommon.History.ForceSave=true;api.CheckChangedDocument();
    }
    return output;
   })().catch(function(error){job.fail(error);throw error;}).finally(function(){saving=null;if(options.onAfterSave)options.onAfterSave();});
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
   if(saving||exporting||openingBusy())throw new Error('A local conversion is already running');
   var job=begin(target==='pdf'?'pdf-export':'export');
   exporting=(async function(){
    if(options.onBeforeSave)options.onBeforeSave();
    if(options.beforeOperationWork)await options.beforeOperationWork();
    if(disposed||(signal()&&signal().aborted))throw new Error('Conversion cancelled');
    var state,output;
    if(target==='pdf'){
     // Share the same insertion barrier as Save, recovery and D's local-picture
     // bridge before the merged document is laid out for the client PDF bundle.
     await pendingImage;
     if(!options.buildPdf)throw new Error('The bundled PDF exporter is unavailable');
     await passwords.confirmExport();output=await options.buildPdf(api,format,downloadOptions||{},signal());
    }else{
     state=await snapshot(downloadOptions&&downloadOptions.oDocumentMailMerge);
     var conversion={resources:state.resources,onProgress:function(value){progress(job,value);},signal:signal()},advanced=downloadOptions&&downloadOptions.advancedOptions;
     if(target==='csv'&&advanced&&advanced.asc_getCodePage){
      var delimiter=advanced.asc_getDelimiter(),custom=advanced.asc_getDelimiterChar();
      if(Array.isArray(delimiter)){
       if(delimiter.length>1)throw new Error('CSV export requires one delimiter');
       delimiter=delimiter[0];
      }
      if(delimiter===-1&&custom)delimiter=0;
      conversion.csv={encoding:advanced.asc_getCodePage(),delimiter:delimiter,delimiterChar:custom};
     }
     output=await passwords.copyAs(target,async function(passwordOptions,protection){return {bytes:await window.EOWasmConvert.convert(state.bytes,state.format,target,Object.assign({},conversion,passwordOptions,{onProgress:function(value){progress(job,Object.assign({},value,passwordOptions.savePassword!==undefined&&value.stage==='converting'?{purpose:'encrypting'}:{}));}})),protection:protection};});
    }
    if(disposed||(signal()&&signal().aborted))throw new Error('Conversion cancelled');
    job.phase('awaiting-output');
    var acknowledgement=await options.onExport(output.bytes,target,downloadOptions||{},output.stats,output.protection);
    if(downloadOptions&&downloadOptions.oDocumentMailMerge)api.eoLastMailMergeExport=acknowledgement&&acknowledgement.kind||'saved';
    job.finish(acknowledgement);
    // A download is a copy: keep the editing format, original and undo/saved state.
    return output.bytes;
   })().catch(function(error){job.fail(error);throw error;}).finally(function(){exporting=null;if(options.onAfterSave)options.onAfterSave();});
   return exporting;
  }
  async function importRecipients(action){
   var job;
   try{
    if(saving||exporting||openingBusy())throw new Error('A local conversion is already running');
    var source=api.mailMergeFileData;job=begin('mail-merge-recipients');
    api.mailMergeFileData=null;
    if(!source||!source.data)throw new Error('Choose a local recipient file');
    if(options.beforeOperationWork)await options.beforeOperationWork();
    var rows=await MailMerge.recipients(source.data,source.fileType,window.EOWasmConvert,{signal:signal(),onProgress:function(value){progress(job,value);}});
    if(disposed||signal()&&signal().aborted)throw new Error('Recipient import cancelled');
    api.asc_StartMailMergeByList(rows);job.finish({kind:'loaded',records:Math.max(0,rows.length-1)});
   }catch(error){if(job)job.fail(error);throw error;}
   finally{if(action)api.sync_EndAction(Asc.c_oAscAsyncActionType.BlockInteraction,action);}
  }
  api.downloadAs=function(action,downloadOptions){
   if(!options.onExport){
    if(downloadOptions&&downloadOptions.fileType!==window.EOWasmConvert.formats[format]){
     report(new Error('Use the desktop Save As command to export this format'));return;
    }
    save().catch(report);return;
   }
   downloadOptions=downloadOptions||{};
   if(action===Asc.c_oAscAsyncAction.MailMergeLoadFile&&downloadOptions.fileType===Asc.c_oAscFileType.JSON){
    api.sync_StartAction(Asc.c_oAscAsyncActionType.BlockInteraction,action);importRecipients(action).catch(report);return;
   }
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
  var localImageUrl=function(name,url){return add.call(this,name,resolveImportedImage(imageUrls,name,url));};
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
  var downloadMailMerge=api.asc_DownloadAsMailMerge;
  var startMailMerge=api.asc_StartMailMerge;
  if(editorFormat==='doct'&&startMailMerge&&options.chooseMailMergeRecipients)api.asc_StartMailMerge=function(source){
   if(source)return startMailMerge.call(api,source);
   Promise.resolve(options.chooseMailMergeRecipients()).then(function(selected){
    if(!selected||disposed)return;
    api.mailMergeFileData={url:undefined,fileType:selected.fileType,data:selected.data};
    api.downloadAs(Asc.c_oAscAsyncAction.MailMergeLoadFile,new Asc.asc_CDownloadOptions(Asc.c_oAscFileType.JSON));
   }).catch(report);
  };
  if(editorFormat==='doct'&&downloadMailMerge)api.asc_DownloadAsMailMerge=function(){
   var args=arguments,document=api.WordControl&&api.WordControl.m_oLogicDocument;
   if(!document||!document.MailMergeMap||!document.MailMergeMap.length)return false;
   api.eoLastMailMergeExport='pending';
   return MailMerge.afterPending(pendingImage,function(){downloadMailMerge.apply(api,args);},report,function(){return !disposed;});
  };
  api.asc_addImage=function(imageOptions){
   var picker=document.createElement('input');picker.type='file';picker.accept='image/png,image/jpeg,image/gif,image/bmp,image/svg+xml,image/webp';
   picker.addEventListener('change',function(){if(picker.files[0]){addImage(picker.files[0],imageOptions).catch(report);}},{once:true});
   picker.click();
  };
  async function open(input,externalJob){
   if(openingBusy())throw Object.assign(new Error('A local operation is already running'),{code:'operation-busy'});
   var job=externalJob||begin('open');activeJob=job;externalOpeningJob=externalJob;opening=true;
   try{
   if(options.beforeOperationWork)await options.beforeOperationWork();
   if(disposed||signal()&&signal().aborted)throw new Error('Conversion cancelled');
   var source=new Uint8Array(input);
   var output=await passwords.open(source,function(passwordOptions){return window.EOWasmConvert.convert(source,format,editorFormat,Object.assign({onProgress:function(value){progress(job,Object.assign({},value,passwordOptions.openPassword!==undefined&&value.stage==='converting'?{purpose:'decrypting'}:{}));},signal:signal()},passwordOptions));});
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
   if(!externalJob)job.finish({kind:'opened'});return output.stats;
   }catch(error){job.fail(error);throw error;}finally{opening=false;}
  }
  async function exportCopy(target,write){
   if(outputs.indexOf(target)<0||target==='pdf')throw new Error('Unsupported local converted copy');
   if(saving||exporting||openingBusy())throw new Error('A local conversion is already running');
   var job=begin('export');
   exporting=(async function(){if(options.beforeOperationWork)await options.beforeOperationWork();if(disposed||(signal()&&signal().aborted))throw new Error('Conversion cancelled');var state=await snapshot();return passwords.copyAs(target,async function(passwordOptions,protection){
    var output=await window.EOWasmConvert.convert(state.bytes,state.format,target,Object.assign({resources:state.resources,signal:signal(),onProgress:function(value){progress(job,Object.assign({},value,passwordOptions.savePassword!==undefined&&value.stage==='converting'?{purpose:'encrypting'}:{}));}},passwordOptions));
    if(disposed||(signal()&&signal().aborted))throw new Error('Conversion cancelled');
    job.phase('awaiting-output');var acknowledgement=await write(output,protection);job.finish(acknowledgement);return output;
   });})().catch(function(error){job.fail(error);throw error;});
   try{return await exporting;}finally{exporting=null;}
  }
  async function saveAs(){
   if(saving||exporting||openingBusy()||operations.isBusy())throw Object.assign(new Error('A local operation is already running'),{code:'operation-busy'});
   var job=begin('save-as'),started=false;
   try{
    if(options.onBeforeSave)options.onBeforeSave();
    if(options.beforeOperationWork)await options.beforeOperationWork();
    return await passwords.saveAs(function(choice){started=true;return save(choice,job,true);});
   }catch(error){job.fail(error);throw error;}
   finally{if(!started&&options.onAfterSave)options.onAfterSave();}
  }
  // Recovery takes a detached snapshot without saving or resetting document history.
  async function snapshot(logicDocument){
   await pendingImage;
   return {bytes:serialize(logicDocument),format:editorFormat,resources:resources.map(function(resource){return {path:resource.path,bytes:new Uint8Array(resource.bytes)};})};
  }
  return {beginOperation:begin,open:open,save:save,exportCopy:exportCopy,saveAs:saveAs,confirmUnprotectedExport:passwords.confirmExport,isProtected:passwords.isProtected,isRecoveryLocked:passwords.isRecoveryLocked,exportAs:exportAs,snapshot:snapshot,serialize:serialize,addImage:addImage,isModified:function(){return api.isDocumentModified();},hasPendingImages:function(){return pendingImages>0;},dispose:function(){disposed=true;if(externalOpeningJob)externalOpeningJob.cancel();operations.dispose();passwords.dispose();pictures.dispose();if(startMailMerge)api.asc_StartMailMerge=startMailMerge;if(downloadMailMerge)api.asc_DownloadAsMailMerge=downloadMailMerge;api.asc_unregisterCallback('asc_onDocumentModifiedChanged',modified);actions.forEach(function(event){api.asc_unregisterCallback(event,action);});urls.forEach(URL.revokeObjectURL);if(docUrls.addImageUrl===localImageUrl)docUrls.addImageUrl=add;}};
 }
 window.Common=window.Common||{};window.Common.ElixirLocal={prepare:prepare,install:install,fileName:FileNames.resolve,resolveImportedImage:resolveImportedImage};
 return window.Common.ElixirLocal;
});

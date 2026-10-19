/*
 * SPDX-FileCopyrightText: 2026 ElixirOffice contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 */
define(['common/main/lib/elixir-local/PdfWriter','common/main/lib/elixir-local/PdfMessages','common/main/lib/elixir-local/Operation'],function(PdfWriter,PdfMessages,Operation){
 'use strict';
 function install(api,options){
  if(!api.DocInfo||!api.DocInfo.get_OfflineApp())throw Error('Local PDF requires an offline editor');
  var opening=false,disposed=false,original=null,saving=null,policy={writable:false,reason:'unverified'},externalOpeningJob;
  var operations=Operation.create(options.onOperation,options.operationScope);
  function openingBusy(){return opening||externalOpeningJob&&externalOpeningJob.isRunning&&externalOpeningJob.isRunning();}
  // Rotation and reordering retain page identity/count; insertion/deletion stay unavailable.
  api.eoLocalPdfCanRotate=function(){return !disposed&&!openingBusy()&&!operations.isBusy()&&!saving&&policy.writable;};
  // Older preservation runtimes cannot attest safe inherited-page reordering.
  api.eoLocalPdfCanReorder=function(){return !disposed&&!openingBusy()&&!operations.isBusy()&&!saving&&policy.writable&&policy.reorderSafe===true;};
  function begin(kind){return operations.begin(kind,{name:api.documentTitle||options.name||''});}
  // Untrusted document scripts and form URI actions must not escape the local transport.
  // This iframe contains only this explicitly offline API instance; hosted SDKs are unchanged.
  function disableScripts(){
   if(!window.AscPDF||!AscPDF.CActionRunScript)throw Error('Local PDF script isolation is unavailable');
   AscPDF.CActionRunScript.prototype.RunScript=function(){};
   AscPDF.CActionRunScript.prototype.Do=function(){var queue=api.getPDFDoc().GetActionsQueue();queue.SetCurAction(this);queue.Continue();};
  }
  var openDocument=api.openDocument;
  api.openDocument=function(){disableScripts();return openDocument.apply(this,arguments);};
  api.asc_registerCallback('asc_onOpenLinkPdfForm',function(_uri,_open,next){if(typeof next==='function')next();});

  function report(error){if(options.onError)options.onError(error);}
  function readOnlyMessage(){return PdfMessages.readOnlyMessage(policy.reason,options.lang,options.pdfStrings);}
  function writerOptions(job){return {signal:signal(),lang:options.lang,strings:options.pdfStrings,onProgress:job&&job.progress};}
  function viewer(){return api.DocumentRenderer||api.getDocumentRenderer();}
  function bytes(){
   var v=viewer(),pages=v.file.pages,originalCount=v.file.originalPagesCount,next=originalCount;
   var indexes=pages.map(function(page){return page.originIndex===undefined?next++:page.originIndex;});
   v.getPDFDoc().BlurActiveObject();v.getPDFDoc().RecalculateAll();
   var output=v.file.nativeFile.SplitPages(indexes,v.SaveForSplit());
   if(!output||output.length>512*1024*1024||output.length<5||String.fromCharCode.apply(null,output.subarray(0,5))!=='%PDF-')throw Error('Local PDF writer did not produce a PDF');
   return new Uint8Array(output);
  }
  // The matched runtime and bounded public preservation gate are verified before this hook.
  function signal(){return options.getSignal&&options.getSignal();}
  async function preserved(job){return PdfWriter.write(original,viewer().Save(),writerOptions(job));}
  async function save(){
   if(openingBusy())throw Object.assign(Error('A local operation is already running'),{code:'operation-busy'});
   if(!policy.writable)throw Error(readOnlyMessage());
   if(disposed)throw Error('The PDF editor is closed');
   if(saving)return saving;
   var job=begin('pdf-save');
   saving=(async function(){
    try{
     if(options.onBeforeSave)options.onBeforeSave();
     if(options.beforeOperationWork)await options.beforeOperationWork();
     if(disposed||signal()&&signal().aborted)throw Error('PDF saving cancelled');
     var changes=viewer().Save(),history=api.getPDFDoc().History,index=history.Index;
     var output=await PdfWriter.write(original,changes,writerOptions(job));
     if(disposed||signal()&&signal().aborted)throw Error('PDF saving cancelled');
     if(!options.onSave)throw Error('No local PDF save destination is available');
     job.phase('awaiting-output');var acknowledgement=await options.onSave(output);
     if(acknowledgement&&acknowledgement.kind==='cancelled'){var cancelled=Error('Save cancelled');cancelled.code='cancelled';throw cancelled;}
     job.finish(acknowledgement);
     // Browser downloads do not acknowledge overwriting the user's source.
     if(options.saveAcknowledgesOriginal===true && history.Index===index){
      history.Reset_SavedIndex(true);api.SetDocumentModified(false);
     }
     return output;
    }catch(error){job.fail(error);throw error;}finally{if(options.onAfterSave)options.onAfterSave();}
   })();
   try{return await saving;}finally{saving=null;}
  }
  api.elixirPdfSaveAllowed=function(){return policy.writable;};
  api.asc_Save=function(isAutoSave){if(isAutoSave===true)return false;if(!policy.writable){report(Error(readOnlyMessage()));return false;}save().catch(report);return true;};
  async function exportAs(target,downloadOptions){
   if(openingBusy())throw Object.assign(Error('A local operation is already running'),{code:'operation-busy'});
   if(target!=='pdf')throw Error('Only PDF copies are available in local PDF mode');
   if(!options.onExport)throw Error('Use the desktop PDF export command');
   var job=begin(downloadOptions&&downloadOptions.isPdfPrint?'print':'pdf-export');
   try{
    if(options.onBeforeSave)options.onBeforeSave();
     if(options.beforeOperationWork)await options.beforeOperationWork();
     if(disposed||signal()&&signal().aborted)throw Error('PDF saving cancelled');
    var output=policy.writable?await preserved(job):bytes();job.phase('awaiting-output');
    var acknowledgement=await options.onExport(output,'pdf',downloadOptions||{},{backend:policy.writable?'original-append':'drawingfile'});
    job.finish(acknowledgement);return output;
   }catch(error){job.fail(error);throw error;}
   finally{if(options.onAfterSave)options.onAfterSave();}
  }
  api.asc_DownloadAs=function(downloadOptions){if(downloadOptions.fileType!==Asc.c_oAscFileType.PDF){report(Error('Only PDF copies are available in local PDF mode'));return false;}exportAs('pdf',downloadOptions).catch(report);return true;};
  api.asc_Print=function(downloadOptions){exportAs('pdf',Object.assign({},downloadOptions,{isPdfPrint:true})).catch(report);return true;};
  async function open(input,externalJob){
   if(openingBusy()||disposed)throw Error('Local PDF opening is unavailable');var job=externalJob||begin('open');externalOpeningJob=externalJob;opening=true;
   try{
   if(options.beforeOperationWork)await options.beforeOperationWork();
   original=new Uint8Array(input);
   if(original.length>512*1024*1024){opening=false;throw Error('Choose a PDF under 512 MiB');}
   if(String.fromCharCode.apply(null,original.subarray(0,5))!=='%PDF-'){opening=false;throw Error('Invalid PDF input');}
   job.phase('initializing');await new Promise(function(resolve,reject){
    var timer=setTimeout(function(){finish(Error('Local PDF did not become ready'));},90000),signal=options.getSignal&&options.getSignal();
    function finish(error){clearTimeout(timer);api.asc_unregisterCallback('asc_onDocumentContentReady',ready);if(signal)signal.removeEventListener('abort',cancel);error?reject(error):resolve();}
    function ready(){
     var v=viewer(),file=v.file,selectAll=file.selectAll;
     file.selectAll=function(){
      if(!v.isFullText){v.fullTextMessageCallback=function(){file.selectAll();};v.showTextMessage();return;}
      selectAll.apply(this,arguments);
      if(this.Selection.Glyph1===-2&&this.Selection.Glyph2===-1){
       this.Selection.IsSelection=true;this.onUpdateSelection();this.onUpdateOverlay();
      }
     };
     finish();
    }
    function cancel(){finish(Error('Opening cancelled'));}
    api.asc_registerCallback('asc_onDocumentContentReady',ready);if(signal)signal.addEventListener('abort',cancel,{once:true});
    if(signal&&signal.aborted)return cancel();
    try{api.asc_openDocumentFromBytes(original);}catch(error){finish(error);}
   });
   try{policy=await PdfWriter.inspect(original,writerOptions(job));}
   catch(error){if(signal()&&signal().aborted)throw error;policy={writable:false,reason:'runtime-unavailable'};}
   if(!policy.writable)api.asc_addRestriction(Asc.c_oAscRestrictionType.View);
   if(!externalJob)job.finish({kind:'opened'});
   return {backend:'drawingfile',inputBytes:original.length,pages:viewer().file.pages.length,policy:policy};}
   catch(error){job.fail(error);throw error;}finally{opening=false;}
  }
  return {beginOperation:begin,open:open,serialize:function(job){return policy.writable?preserved(job):bytes();},exportAs:exportAs,save:save,readOnlyMessage:readOnlyMessage,canSave:function(){return policy.writable;},isSaving:function(){return !!saving;},snapshot:async function(){return {bytes:policy.writable?await preserved():original.slice(),format:'pdf',resources:[]};},isModified:function(){return api.isDocumentModified();},hasPendingImages:function(){return false;},dispose:function(){disposed=true;if(externalOpeningJob)externalOpeningJob.cancel();operations.dispose();}};
 }
 return {install:install};
});

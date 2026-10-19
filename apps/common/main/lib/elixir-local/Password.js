/* SPDX-License-Identifier: AGPL-3.0-only */
// Load dialog views on demand after the editor core UI starts. Eager imports
// would evaluate Common.UI before the application bootstrap has created it.
define(['require'],function(require){
 'use strict';
 var editors=new WeakMap();
 function dialogs(){return new Promise(function(resolve,reject){require(['common/main/lib/view/OpenDialog','common/main/lib/view/PasswordDialog'],resolve,reject);});}
 function error(code,text){var e=new Error(text);e.code=code;return e;}
 function text(key,fallback){
  var locale=Common.Locale,value;
  // The locale or lazy dialog view may not be ready when a safety warning opens.
  try{value=locale&&typeof locale.get==='function'?locale.get(key,{name:'Common.Views.PasswordDialog',default:fallback}):fallback;}catch(_error){return fallback;}
  return typeof value==='string'&&value.trim()?value:fallback;
 }
 function prepare(api,main){
  if(!api.eoLocalMode)return;
  editors.set(api,main);
  api.asc_isProtectionSupport=function(){return !!(api.eoLocalMode&&api.DocInfo&&api.DocInfo.get_OfflineApp()&&['docx','xlsx','pptx','odt','ods','odp'].indexOf(api.documentFormat)>=0&&main.appOptions.isEdit&&(!main.permissions||main.permissions.protect!==false));};
 }
 function install(api,options){
  if(!api.eoLocalMode||!api.DocInfo||!api.DocInfo.get_OfflineApp())throw error('permission','Password integration requires a local offline editor');
  var activePassword,locked=false,lockConfirmed=false,disposed=false,dialog=null,cancelPrompt=null,changing=false;
  var original={sendEvent:api.sendEvent,setPassword:api.asc_setCurrentPassword,resetPassword:api.asc_resetPassword};
  function phase(stage){try{if(options.onState){var value=options.onState({stage:stage});if(value&&value.catch)value.catch(function(){});}}catch(_error){}}
  function signal(){return options.getSignal?options.getSignal():options.signal;}
  function alive(){if(disposed||(signal()&&signal().aborted))throw error('cancelled','Local password operation cancelled');}
  function protectedDocument(){return activePassword!==undefined;}
  function notify(){original.sendEvent.call(api,'asc_onDocumentPassword',protectedDocument());}
  function permissions(){return api.asc_isProtectionSupport()&&!api.isViewMode;}
  async function lock(){
   locked=true;if(lockConfirmed)return;
   if(options.onProtectionLock)await options.onProtectionLock();
   lockConfirmed=true;
  }
  function prompt(build){
   alive();phase('awaiting-user');return new Promise(function(resolve,reject){
    var done=false,currentSignal=signal();
    function cancel(){finish(error('cancelled','Local password operation cancelled'));}
    function finish(e,value){
     if(done)return;done=true;
     if(currentSignal)currentSignal.removeEventListener('abort',cancel);
     var current=dialog;dialog=null;cancelPrompt=null;
     if(current&&current.inputPwd)current.inputPwd.setValue('');
     if(current&&current.repeatPwd)current.repeatPwd.setValue('');
     if(current)current.close();
     if(!e)phase('preparing');e?reject(e):resolve(value);
    }
    cancelPrompt=cancel;dialog=build(finish);
    if(currentSignal)currentSignal.addEventListener('abort',cancel,{once:true});
    if(currentSignal&&currentSignal.aborted)cancel();
   });
  }
  async function openPassword(incorrect){
   await dialogs();alive();
   var main=editors.get(api);
   // Match the normal encrypted-open dialog: the document load mask must not
   // intercept pointer and keyboard input while a password is requested.
   if(main){
    if(main.loadMask)main.loadMask.hide();
    if(main.onLongActionEnd)main.onLongActionEnd(Asc.c_oAscAsyncActionType.BlockInteraction,Common.UI.blockOperations.LoadingDocument);
   }
   return prompt(function(finish){
   var view=new Common.Views.OpenDialog({title:Common.Views.OpenDialog.prototype.txtTitleProtected,
    type:Common.Utils.importTextType.DRM,closable:true,warning:false,validatePwd:incorrect,maxPasswordLength:255,
    handler:function(result,value){if(result!=='ok')return finish(error('cancelled','Opening cancelled'));var password=value.drmOptions.asc_getPassword();try{window.EOWasmConvert.validatePassword(password);finish(null,password);}catch(e){finish(e);}}
   });view.show();return view;
  });}
  async function newPassword(){await dialogs();alive();return prompt(function(finish){
   var labels={};['txtTitle','txtPassword','txtDescription','txtRepeat','txtIncorrectPwd','txtWarning'].forEach(function(key){labels[key]=text(key,Common.Views.PasswordDialog.prototype[key]);});
   var view=new Common.Views.PasswordDialog({api:api,cls:'modal-dlg eo-local-password',localPasswordLabels:labels,handler:function(result,value){if(result!=='ok')return finish(error('cancelled','Password change cancelled'));try{window.EOWasmConvert.validatePassword(value);finish(null,value);}catch(e){finish(e);}}});
   view.show();return view;
  });}
  async function choose(){
   await dialogs();alive();
   var selected=await prompt(function(finish){return Common.UI.alert({
    title:text('txtPassword','Password'),msg:text('txtDescription','A Password is required to open this document'),
    width:Math.min(780,Math.max(240,(window.innerWidth||800)-32)),footerCls:'center eo-local-password-actions',
    buttons:[{value:'keep',caption:protectedDocument()?text('txtLocalKeepCurrentPassword','Keep current password'):text('txtLocalKeepWithoutPassword','Keep without password')},
     {value:'change',caption:protectedDocument()?text('txtLocalChangePassword','Change password'):text('txtLocalAddPassword','Add password')},
     {value:'remove',caption:text('txtLocalRemovePassword','Remove password')},'cancel'],
    callback:function(value){if(value==='keep'||value==='change'||value==='remove')finish(null,value);else finish(error('cancelled','Save As cancelled'));}
   });});
   return selected==='change'?newPassword():selected==='remove'?null:undefined;
  }
  async function confirmExport(){
   if(!protectedDocument())return false;
   var accepted=await prompt(function(finish){return Common.UI.warning({
    msg:text('txtLocalPlaintextCopyWarning','This exported copy will not be password protected. Continue?'),buttons:['no','yes'],
    callback:function(value){finish(null,value==='yes');}
   });});
   if(!accepted)throw error('cancelled','Export cancelled');return true;
  }
  async function open(input,convert){
   var password,incorrect=false;
   var compound=input.length>=8&&[208,207,17,224,161,177,26,225].every(function(v,i){return input[i]===v;});
   if(compound)await lock();
   while(true){
    alive();
    try{
     var output=await convert(password===undefined?{}:{openPassword:password});alive();
     if(compound&&password===undefined)throw error('unsupported-encryption','An encrypted Office file without an explicit nonempty password cannot be edited locally');
     activePassword=password;notify();return output;
    }catch(e){
     if(e.code!=='password-required'&&e.code!=='password-incorrect')throw e;
     await lock();password=await openPassword(incorrect||e.code==='password-incorrect');incorrect=true;
    }
   }
  }
  async function save(convertAndWrite,choice){
   alive();if(choice!==undefined&&choice!==null)window.EOWasmConvert.validatePassword(choice);
   var candidate=choice===undefined?activePassword:choice===null?undefined:choice;
   if(candidate!==undefined||protectedDocument())await lock();
   var removed=protectedDocument()&&candidate===undefined;
   if(removed){
    var accepted=await prompt(function(finish){return Common.UI.warning({
     msg:text('txtLocalRemovePasswordWarning','This save will remove password protection and may overwrite the encrypted original with an unprotected file. Continue?'),buttons:['no','yes'],
     callback:function(value){finish(null,value==='yes');}
    });});
    if(!accepted)throw error('cancelled','Password removal cancelled');
    alive();
   }
   var result=await convertAndWrite(candidate===undefined?{}:{savePassword:candidate},{protected:candidate!==undefined,remove:removed});
   if(!disposed){activePassword=candidate;notify();}return result;
  }
  async function copy(target,convertAndWrite){
   alive();var supported=['docx','xlsx','pptx','odt','ods','odp'].indexOf(target)>=0;
   if(!supported){await confirmExport();return convertAndWrite({},{protected:false,remove:protectedDocument()});}
   return convertAndWrite(activePassword===undefined?{}:{savePassword:activePassword});
  }
  async function copyAs(target,convertAndWrite){
   if(!protectedDocument())return convertAndWrite({},{protected:false,remove:false});
   if(['docx','xlsx','pptx','odt','ods','odp'].indexOf(target)<0)return copy(target,convertAndWrite);
   var choice=await choose(),candidate=choice===undefined?activePassword:choice===null?undefined:choice;
   if(candidate!==undefined)await lock();
   return convertAndWrite(candidate===undefined?{}:{savePassword:candidate},{protected:candidate!==undefined,remove:protectedDocument()&&candidate===undefined});
  }
  function change(value){
   if(changing||options.isBusy())return Promise.reject(error('busy','A local save is already running'));
   if(!permissions())return Promise.reject(error('permission','Password protection requires local editing permission'));
   changing=true;return Promise.resolve().then(function(){return options.save(value);}).finally(function(){changing=false;});
  }
  api.sendEvent=function(name){if(name==='asc_onDocumentPassword'){var args=Array.prototype.slice.call(arguments);args[1]=protectedDocument();return original.sendEvent.apply(this,args);}return original.sendEvent.apply(this,arguments);};
  api.asc_setCurrentPassword=function(value){change(value).catch(options.onError);};
  api.asc_resetPassword=function(){change(null).catch(options.onError);};
  return {open:open,save:save,copy:copy,copyAs:copyAs,confirmExport:confirmExport,isProtected:protectedDocument,isRecoveryLocked:function(){return locked;},
   saveAs:async function(run){if(changing||options.isBusy())throw error('busy','A local save is already running');return run(protectedDocument()?await choose():undefined);},
   dispose:function(){disposed=true;activePassword=undefined;if(cancelPrompt)cancelPrompt();if(dialog){dialog.close();dialog=null;}api.sendEvent=original.sendEvent;api.asc_setCurrentPassword=original.setPassword;api.asc_resetPassword=original.resetPassword;api.currentPassword='';api.currentPasswordOld='';}
  };
 }
 return {prepare:prepare,install:install};
});

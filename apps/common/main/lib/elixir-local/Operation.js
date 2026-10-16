/*
 * SPDX-FileCopyrightText: 2026 ElixirOffice contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 */
(function(factory){
 if(typeof module==='object'&&module.exports)module.exports=factory();
 else if(typeof define==='function'&&define.amd)define([],factory);
 else {var common=globalThis.Common=globalThis.Common||{};common.ElixirLocalOperation=factory();}
})(function(){
 'use strict';
 var sequence=0,phases=['preparing','loading-runtime','initializing','converting','awaiting-user','awaiting-output','writing'];
 function create(notify,scope,idPrefix){
  idPrefix=/^[a-z-]{1,32}$/.test(idPrefix||'')?idPrefix:'local-operation';
  var current=null,disposed=false;
  function publish(job,phase,detail){
   var event=Object.assign({id:job.id,kind:job.kind,scope:scope,phase:phase},job.name===undefined?{}:{name:job.name},detail||{});
   try{if(notify){var result=notify(event);if(result&&typeof result.catch==='function')result.catch(function(){});}}catch(_error){/* An observer must never turn a committed save into a failure. */}
   return true;
  }
  function emit(job,phase,detail){if(job.terminal||current!==job)return false;return publish(job,phase,detail);}
  function begin(kind,detail){
   if(disposed)throw Error('The local operation tracker is closed');
   if(current&&!current.terminal)throw Object.assign(Error('A local operation is already running'),{code:'operation-busy'});
   var job={id:idPrefix+'-'+(++sequence),kind:kind,terminal:false,name:detail&&typeof detail.name==='string'?detail.name.slice(0,512):undefined};current=job;
   emit(job,'preparing');
   function terminal(phase,detail){
    if(job.terminal||current!==job)return false;job.terminal=true;return publish(job,phase,detail);
   }
   return {
    id:job.id,
    isRunning:function(){return !job.terminal&&current===job;},
    progress:function(value){
     if(!value||phases.indexOf(value.stage)<0)return false;
     var details={};
     if(typeof value.cancellable==='boolean')details.cancellable=value.cancellable;
     if(['decrypting','encrypting'].indexOf(value.purpose)>=0)details.purpose=value.purpose;
     if(Number.isInteger(value.completed)&&Number.isInteger(value.total)&&value.total>0&&value.completed>=0&&value.completed<=value.total){
      details.completed=value.completed;details.total=value.total;
     }
     return emit(job,value.stage,details);
    },
    phase:function(stage){return phases.indexOf(stage)>=0&&emit(job,stage);},
    finish:function(result){
     if(result&&result.kind==='cancelled')return terminal('cancelled');
     if(result&&result.kind==='opened')return terminal('complete',{outcome:'opened'});
     if(result&&result.kind==='saved'&&result.committed===true){
      return terminal(result.durabilityUnconfirmed?'warning':'ready',{outcome:'saved',name:result.name||null,artifactId:result.artifactId||null,bytes:result.bytes,reason:result.durabilityUnconfirmed?'durability-unconfirmed':null});
     }
     if(result&&['download-started','print-started','printed'].indexOf(result.kind)>=0){
      return terminal('complete',{outcome:result.kind,name:result.name||null,bytes:result.bytes});
     }
     // Resolving a conversion promise alone is not a durable-file acknowledgement.
     return terminal('complete',{outcome:'completion-unconfirmed'});
    },
    fail:function(error){
     var cancelled=error&&(error.code==='cancelled'||error.code==='aborted'||error.name==='AbortError'||/^(?:Conversion cancelled|PDF saving cancelled|Save cancelled)(?::|$)/.test(error.message||''));
     return terminal(cancelled?'cancelled':'error',cancelled?{}:{reason:error&&error.message||'Local operation failed'});
    },
    cancel:function(){return terminal('cancelled');}
   };
  }
  return {begin:begin,isBusy:function(){return !!current&&!current.terminal;},
   dispose:function(){disposed=true;if(current&&!current.terminal){current.terminal=true;publish(current,'cancelled',{outcome:'disposed'});}current=null;}};
 }
 return {create:create};
});

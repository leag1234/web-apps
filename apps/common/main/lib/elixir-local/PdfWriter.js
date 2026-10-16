/*
 * SPDX-FileCopyrightText: 2026 ElixirOffice contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 */
define(['common/main/lib/elixir-local/PdfMessages'],function(PdfMessages){
 'use strict';
 var base=new URL('../../../../sdkjs/common/clientpdf/runtime/',window.location.href).href;
 function workerMain(){
  self.onmessage=async function(event){
   try{
    var job=event.data;
    function stage(value){if(job.progress===true)self.postMessage({type:'progress',stage:value});}
    stage('loading-runtime');
    var current=await fetch(job.base+'current.json',{cache:'no-store',credentials:'same-origin'});
    if(!current.ok)throw Error('PDF preservation runtime is unavailable');
    var pin=await current.json();
    if(!/^[a-f0-9]{12}-[a-f0-9]{12}-[a-z0-9-]+$/.test(pin.artifact_version))throw Error('Invalid PDF runtime pin');
    var runtime=job.base+pin.artifact_version+'/';
    var response=await fetch(runtime+'manifest.json',{cache:'force-cache',credentials:'same-origin'});
    if(!response.ok)throw Error('PDF preservation manifest is unavailable');
    var manifest=await response.json();
    if(manifest.artifact_version!==pin.artifact_version||manifest.build_flags.module_kind!=='pdf'||manifest.build_flags.pdf_original_append!==true)throw Error('PDF preservation is not supported by this runtime');
    async function checked(name){
     var item=manifest.artifacts[name],response=await fetch(runtime+name,{cache:'force-cache',credentials:'same-origin'});
     if(!item||!response.ok)throw Error('PDF runtime artifact is unavailable');
     var bytes=new Uint8Array(await response.arrayBuffer());
     var digest=new Uint8Array(await crypto.subtle.digest('SHA-256',bytes));
     var hash=Array.from(digest,function(value){return value.toString(16).padStart(2,'0');}).join('');
     if(bytes.length!==item.bytes||hash!==item.sha256)throw Error('PDF runtime checksum mismatch');
     return bytes;
    }
    await checked('x2t.js');var wasm=await checked('x2t.wasm');
    importScripts(runtime+'x2t.js');
    stage('initializing');
    var mod=await createX2t({noInitialRun:true,wasmBinary:wasm,print:function(){},printErr:function(){}});
    var FS=mod.FS;
    ['/p','/fonts','/tmp/pdf'].forEach(function(dir){FS.mkdirTree(dir);});
    FS.writeFile('/p/in.pdf',job.original);
    if(job.changes)FS.writeFile('/p/changes.bin',job.changes);
    job.fonts.forEach(function(bytes,index){FS.writeFile('/fonts/font'+index+'.ttf',bytes);});
    var exit;
    stage('converting');
    try{exit=mod.callMain([job.inspect?'--inspect-pdf':'--edit-pdf']);}
    catch(error){if(typeof error.status==='number')exit=error.status;else throw error;}
    var policy=JSON.parse(FS.readFile('/p/policy.json',{encoding:'utf8'}));
    if(typeof policy.writable!=='boolean'||policy.writer!=='original-append')throw Error('Invalid PDF preservation policy');
    if(job.inspect){self.postMessage({policy:policy});return;}
    if(exit!==0||!policy.writable)throw Object.assign(Error('This PDF cannot be saved safely. Save a copy instead; the original stays unchanged.'),{reason:policy.reason||'writer-failure'});
    var output=FS.readFile('/p/out.pdf');
    if(output.length<job.original.length||output.length>256*1024*1024)throw Error('Invalid preserved PDF output');
    for(var i=0;i<job.original.length;i++)if(output[i]!==job.original[i])throw Error('PDF original-byte preservation failed');
    self.postMessage({policy:policy,bytes:output},[output.buffer]);
   }catch(error){self.postMessage({error:String(error.message||error),reason:error.reason||'runtime-unavailable'});}
  };
 }
 function fonts(){
  var result=[],size=0,streams=window.AscFonts&&AscFonts.g_fonts_streams||[];
  streams.forEach(function(stream){
   if(!stream)return;
   var length=stream.size===undefined?stream.len:stream.size,data=stream.data,bytes;
   if(typeof data==='number')bytes=AscFonts.GetUint8ArrayFromPointer(data,length).slice();
   else if(data instanceof Uint8Array)bytes=data.subarray(0,length===undefined?data.length:length).slice();
   else if(data instanceof ArrayBuffer)bytes=new Uint8Array(data,0,length===undefined?data.byteLength:length).slice();
   if(!bytes)return;
   size+=bytes.length;if(result.length>=64||size>64*1024*1024)throw Error('PDF font resources exceed the local save limit');
   result.push(bytes);
  });return result;
 }
 function run(original,changes,inspect,options){
  options=options||{};
  if(!original||original.length>128*1024*1024)return Promise.reject(Error('Choose a PDF under 128 MiB for preserved saving'));
  if(changes&&changes.length>64*1024*1024)return Promise.reject(Error('PDF changes exceed the local save limit'));
  return new Promise(function(resolve,reject){
   var url=URL.createObjectURL(new Blob(['('+workerMain.toString()+')();'],{type:'text/javascript'})),worker;
   try{worker=new Worker(url);}catch(error){URL.revokeObjectURL(url);reject(error);return;}
   var signal=options.signal,timer=setTimeout(function(){finish(Object.assign(Error('PDF preservation timed out'),{reason:'timeout'}));},60000),done=false;
   function finish(error,value){if(done)return;done=true;clearTimeout(timer);worker.terminate();URL.revokeObjectURL(url);if(signal)signal.removeEventListener('abort',cancel);error?reject(error):resolve(value);}
   function cancel(){finish(Object.assign(Error('PDF saving cancelled'),{reason:'cancelled',code:'cancelled',name:'AbortError'}));}
   worker.onmessage=function(event){if(done)return;var value=event.data;if(value&&value.type==='progress'){if(['loading-runtime','initializing','converting'].indexOf(value.stage)>=0&&options.onProgress){try{var result=options.onProgress({stage:value.stage});if(result&&result.catch)result.catch(function(){});}catch(_error){}}return;}if(!value||value.error)finish(Object.assign(Error(value&&value.error||'Invalid PDF writer response'),{reason:value&&value.reason||'runtime-unavailable'}));else finish(null,value);};
   worker.onerror=function(event){event.preventDefault();finish(Error('Local PDF preservation failed'));};
   if(signal){signal.addEventListener('abort',cancel,{once:true});if(signal.aborted){cancel();return;}}
   try{worker.postMessage({base:base,original:new Uint8Array(original),changes:changes&&new Uint8Array(changes),inspect:inspect,progress:typeof options.onProgress==='function',fonts:inspect?[]:fonts()});}
   catch(error){finish(error);}
  });
 }
 return {inspect:function(original,options){return run(original,null,true,options).then(function(result){if(!result.policy||typeof result.policy.writable!=='boolean')throw Error('Invalid PDF policy response');return result.policy;});},write:function(original,changes,options){return run(original,changes,false,options).catch(function(error){if(error.reason)error.message=PdfMessages.message(error.reason,options&&options.lang,options&&options.strings);throw error;}).then(function(result){if(!(result.bytes instanceof Uint8Array)||result.bytes.length<5||String.fromCharCode.apply(null,result.bytes.subarray(0,5))!=='%PDF-')throw Error('Invalid preserved PDF response');return result.bytes;});}};
});

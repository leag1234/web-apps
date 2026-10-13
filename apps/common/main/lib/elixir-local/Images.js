/*
 * SPDX-FileCopyrightText: 2026 ElixirOffice contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 */
(function(factory){
 if(typeof define==='function'&&define.amd)define([],factory);
 else {window.Common=window.Common||{};window.Common.ElixirLocalImages=factory();}
})(function(){
 'use strict';
 // Captured before the general offline guard; used only by an explicit URL command.
 var fetchImage=window.fetch.bind(window),owners=new Map(),originalSend,originalUpload;
 var MAX_BYTES=20*1024*1024,MAX_PIXELS=16*1024*1024,TIMEOUT=15000;
 function failure(message){return new Error(message);}
 function dimensions(bytes){
  var v=new DataView(bytes.buffer,bytes.byteOffset,bytes.byteLength);
  if(bytes.length>=24&&v.getUint32(0)===0x89504e47&&v.getUint32(4)===0x0d0a1a0a)
   return {type:'image/png',extension:'png',width:v.getUint32(16),height:v.getUint32(20)};
  if(bytes.length>=4&&bytes[0]===255&&bytes[1]===216){
   for(var p=2;p+4<bytes.length;){
    if(bytes[p++]!==255)continue;
    while(bytes[p]===255)p++;
    var marker=bytes[p++];
    if(marker===217||marker===218)break;
    if(marker===216||marker===1||marker>=208&&marker<=215)continue;
    var length=v.getUint16(p);
    if(length<2||p+length>bytes.length)break;
    if([192,193,194,195,197,198,199,201,202,203,205,206,207].includes(marker)&&length>=7)
     return {type:'image/jpeg',extension:'jpg',width:v.getUint16(p+5),height:v.getUint16(p+3)};
    p+=length;
   }
  }
  var ascii=function(start,length){return String.fromCharCode.apply(null,bytes.subarray(start,start+length));};
  if(bytes.length>=10&&/^GIF8[79]a$/.test(ascii(0,6)))
   return {type:'image/gif',extension:'png',raster:true,width:v.getUint16(6,true),height:v.getUint16(8,true)};
  if(bytes.length>=26&&ascii(0,2)==='BM'){
   var header=v.getUint32(14,true);
   if(header===12)return {type:'image/bmp',extension:'png',raster:true,width:v.getUint16(18,true),height:v.getUint16(20,true)};
   if(header>=40)return {type:'image/bmp',extension:'png',raster:true,width:Math.abs(v.getInt32(18,true)),height:Math.abs(v.getInt32(22,true))};
  }
  if(bytes.length>=30&&ascii(0,4)==='RIFF'&&ascii(8,4)==='WEBP'){
   var kind=ascii(12,4),width,height;
   if(kind==='VP8X'){width=1+bytes[24]+(bytes[25]<<8)+(bytes[26]<<16);height=1+bytes[27]+(bytes[28]<<8)+(bytes[29]<<16);}
   else if(kind==='VP8 '&&bytes[23]===157&&bytes[24]===1&&bytes[25]===42){width=v.getUint16(26,true)&16383;height=v.getUint16(28,true)&16383;}
   else if(kind==='VP8L'&&bytes[20]===47){var bits=v.getUint32(21,true);width=(bits&16383)+1;height=((bits>>>14)&16383)+1;}
   if(width&&height)return {type:'image/webp',extension:'png',raster:true,width:width,height:height};
  }
  if(bytes.length<1024*1024){
   var text=new TextDecoder().decode(bytes);
   if(/^\s*(?:<\?xml[^>]*>\s*)?<svg[\s>]/i.test(text)){
    if(/<!DOCTYPE|<!ENTITY/i.test(text))throw failure('SVG document types and entities are unavailable. Export it as PNG first.');
    var xml=new DOMParser().parseFromString(text,'image/svg+xml'),root=xml.documentElement;
    if(root.localName!=='svg'||xml.querySelector('parsererror')||xml.doctype)throw failure('The SVG picture is invalid.');
    // No SVG active content or external resources; rasterize before embedding.
    var nodes=xml.querySelectorAll('*');if(nodes.length>2000)throw failure('This SVG is too complex. Export it as PNG first.');
    for(var node of nodes){
     if(['script','foreignObject','image','use','style','a','animate','animateTransform','set','filter'].includes(node.localName))
      throw failure('This SVG uses active content or linked resources. Export it as PNG first.');
     for(var attribute of node.attributes){
      if(node===root&&attribute.name==='style'||/^on/i.test(attribute.name)||/href/i.test(attribute.name)||/url\s*\(|@import/i.test(attribute.value))
       throw failure('This SVG uses linked resources. Export it as PNG first.');
     }
    }
    var box=(root.getAttribute('viewBox')||'').trim().split(/[\s,]+/).map(Number);
    var dimension=function(name,index){
     var value=root.getAttribute(name);if(value&&!/^\d+(?:\.\d+)?(?:px)?$/.test(value))throw failure('Use SVG dimensions in pixels.');
     return value?parseFloat(value):box.length===4?box[index]:0;
    };
    return {type:'image/svg+xml',extension:'png',raster:true,width:dimension('width',2),height:dimension('height',3)};
   }
  }
  throw failure('Local pictures must contain a valid PNG, JPEG, GIF, BMP, SVG or WebP image.');
 }
 function assertAvailable(){if(owners.size)throw failure('Only one local picture session can be active in an editor frame.');}
 function install(api,options){
  assertAvailable();
  var disposed=false,requests=new Set(),key=api,formatOwners=[];
  if(owners.has(key))throw failure('A local picture session is already active.');
  function active(){if(disposed||options.signal()&&options.signal().aborted)throw failure('Picture insertion cancelled.');}
  function bounded(job){
   active();
   var controller=new AbortController(),parent=options.signal(),cancel=function(){controller.abort();};
   requests.add(controller);if(parent)parent.addEventListener('abort',cancel,{once:true});
   var timer=setTimeout(cancel,TIMEOUT);
   return Promise.race([job(controller.signal),new Promise(function(_,reject){
    controller.signal.addEventListener('abort',function(){reject(failure('Picture insertion cancelled or timed out.'));},{once:true});
   })]).finally(function(){clearTimeout(timer);requests.delete(controller);if(parent)parent.removeEventListener('abort',cancel);});
  }
  async function bytesFromUrl(source,allowRemote){
   var data=/^data:/i.test(source),limit=data?Math.ceil(MAX_BYTES*4/3)+1024:8192;
   if(source.length>limit)throw failure(data?'The encoded picture exceeds the local size limit.':'The picture URL is too long.');
   var url=new URL(source,window.location.href),remote=/^https?:$/.test(url.protocol);
   if(remote&&!allowRemote)throw failure('Remote pictures from pasted content are unavailable offline. Use Insert from URL explicitly.');
   if(!remote&&!['blob:','data:'].includes(url.protocol))throw failure('Picture URLs must use HTTP or HTTPS.');
   if(url.username||url.password)throw failure('Picture URLs cannot contain credentials.');
   return bounded(async function(signal){
    var response;
    try{response=await fetchImage(url.href,{method:'GET',mode:'cors',credentials:'omit',redirect:'error',referrerPolicy:'no-referrer',cache:'no-store',signal:signal});}
    catch(error){throw failure('The picture URL is unavailable in local mode. Check the connection, CORS permission and offline security policy, or insert the picture from a file.');}
    if(!response.ok||response.type==='opaque')throw failure('The picture URL could not be fetched. Check the connection and CORS permission.');
    if(Number(response.headers.get('content-length'))>MAX_BYTES){if(response.body)await response.body.cancel();throw failure('Pictures must be smaller than 20 MiB.');}
    if(!response.body)throw failure('The picture response is empty.');
    var reader=response.body.getReader(),parts=[],length=0;
    try{
     for(;;){var part=await reader.read();if(part.done)break;length+=part.value.length;if(length>MAX_BYTES)throw failure('Pictures must be smaller than 20 MiB.');parts.push(part.value);}
    }catch(e){await reader.cancel().catch(function(){});throw e;}
    var bytes=new Uint8Array(length),offset=0;parts.forEach(function(part){bytes.set(part,offset);offset+=part.length;});return bytes;
   });
  }
  async function importImage(source,allowRemote){
   active();
   if(typeof source==='string'){
    var known=options.lookup(source);if(known)return known;
   }else if(source.size>MAX_BYTES)throw failure('Pictures must be smaller than 20 MiB.');
   var bytes=typeof source==='string'?await bytesFromUrl(source,allowRemote):new Uint8Array(await source.arrayBuffer());
   active();if(!bytes.length||bytes.length>MAX_BYTES)throw failure('Pictures must be smaller than 20 MiB.');
   var info=dimensions(bytes);
   if(!info.width||!info.height||info.width>32768||info.height>32768||info.width*info.height>MAX_PIXELS)
    throw failure('The picture dimensions exceed the local decoding limit.');
   var url=URL.createObjectURL(new Blob([bytes],{type:info.type})),committed=false,decoded;
   try{
    await bounded(function(signal){return new Promise(function(resolve,reject){
     var image=new Image(),cancel=function(){image.src='';finish(failure('Picture insertion cancelled.'));};
     function finish(error){image.onload=image.onerror=null;signal.removeEventListener('abort',cancel);error?reject(error):resolve();}
     image.onload=function(){decoded=image;finish();};image.onerror=function(){finish(failure('The picture could not be decoded.'));};
     signal.addEventListener('abort',cancel,{once:true});image.src=url;
    });});
    active();
    if(info.raster){
     var canvas=document.createElement('canvas');canvas.width=info.width;canvas.height=info.height;
     canvas.getContext('2d').drawImage(decoded,0,0,canvas.width,canvas.height);
     var raster=await bounded(function(){return new Promise(function(resolve,reject){canvas.toBlob(function(blob){blob?resolve(blob):reject(failure('The picture could not be converted to PNG.'));},'image/png');});});
     if(raster.size>MAX_BYTES)throw failure('The converted picture exceeds 20 MiB.');bytes=new Uint8Array(await raster.arrayBuffer());
     URL.revokeObjectURL(url);url=URL.createObjectURL(raster);
    }
    await bounded(function(){return new Promise(function(resolve){api.ImageLoader.LoadImagesWithCallback([url],resolve,[]);});});
    active();
    var result=options.register(bytes,info.extension,url);committed=true;return result;
   }finally{if(!committed)URL.revokeObjectURL(url);}
  }
  function message(error){if(!disposed&&window.Common&&Common.UI&&Common.UI.warning)Common.UI.warning({msg:error.message});}
  function queue(job){return options.enqueue(job);}
  function send(images,callback,allowRemote,quiet){
   var block=!api.isOpenedFrameEditor;
   if(block)api.sync_StartAction(Asc.c_oAscAsyncActionType.BlockInteraction,Asc.c_oAscAsyncAction.LoadImage);
   queue(async function(){
    var results=[];
    try{
     for(var source of images){try{results.push(await importImage(source,allowRemote));}catch(e){results.push({url:'error',path:'error'});if(!quiet)message(e);}}
    }finally{if(block)api.sync_EndAction(Asc.c_oAscAsyncActionType.BlockInteraction,Asc.c_oAscAsyncAction.LoadImage);callback(results);}
   }).catch(function(error){options.onError(error);});
  }
  var previousAdd=api.AddImageUrl;
  var addUrl=function(urls,imgProp,token,obj){
   if(api.isLongAction())return;
   send(Array.from(urls),function(data){var valid=data.filter(function(item){return item.url!=='error';}).map(function(item){return item.url;});if(valid.length&&!disposed)api._addImageUrl(valid,obj);},true,false);
  };
  api.AddImageUrl=addUrl;
  var previousCheck=api.asc_checkImageUrlAndAction;
  var checkUrl=function(source,callback){
   api.sync_StartAction(Asc.c_oAscAsyncActionType.BlockInteraction,Asc.c_oAscAsyncAction.UploadImage);
   queue(async function(){
    try{
     var image=await importImage(source,false);
     if(!disposed)api.asc_loadLocalImageAndAction(AscCommon.g_oDocumentUrls.imagePath2Local(image.path),callback);
    }catch(error){message(error);}
    finally{api.sync_EndAction(Asc.c_oAscAsyncActionType.BlockInteraction,Asc.c_oAscAsyncAction.UploadImage);}
   }).catch(options.onError);
  };
  api.asc_checkImageUrlAndAction=checkUrl;

  var owner={api:api,send:function(images,callback,quiet){send(images,callback,false,quiet);},upload:function(files,callback){
   queue(async function(){
    var error=Asc.c_oAscError.ID.No,urls=[];
    try{if(!files.length)throw failure('No picture was selected.');for(var file of files)urls.push((await importImage(file,false)).url);}
    catch(e){error=Asc.c_oAscError.ID.UplImageUrl;message(e);}
    callback(error,urls);
   }).catch(options.onError);
  }};
  owners.set(key,owner);
  function activate(){
   // The SDK drop validator shares this table within one editor frame.
   var formats=AscCommon.c_oAscImageUploadProp;
   if(formats&&!formats.SupportedFormats.includes('webp')){
    var previous=formats.SupportedFormats,local=previous.concat('webp');
    formatOwners.push({object:formats,previous:previous,local:local});formats.SupportedFormats=local;
   }
   if(api.AddImageUrl!==addUrl){previousAdd=api.AddImageUrl;api.AddImageUrl=addUrl;}
   if(api.asc_checkImageUrlAndAction!==checkUrl){previousCheck=api.asc_checkImageUrlAndAction;api.asc_checkImageUrlAndAction=checkUrl;}
   // Some editor SDK modules finish loading during openDocumentFromBytes.
   if(AscCommon.sendImgUrls!==sendWrapper){originalSend=AscCommon.sendImgUrls;AscCommon.sendImgUrls=sendWrapper;}
   if(AscCommon.UploadImageFiles!==uploadWrapper){originalUpload=AscCommon.UploadImageFiles;AscCommon.UploadImageFiles=uploadWrapper;}
  }
  activate();
  return {activate:activate,insert:function(file,imageOptions){return queue(async function(){var image=await importImage(file,false);if(!disposed)api._addImageUrl([image.url],imageOptions);});},dispose:function(){
   disposed=true;formatOwners.forEach(function(owner){if(owner.object.SupportedFormats===owner.local)owner.object.SupportedFormats=owner.previous;});requests.forEach(function(request){request.abort();});owners.delete(key);
   if(api.AddImageUrl===addUrl)api.AddImageUrl=previousAdd;
   if(api.asc_checkImageUrlAndAction===checkUrl)api.asc_checkImageUrlAndAction=previousCheck;
   if(!owners.size){if(AscCommon.sendImgUrls===sendWrapper)AscCommon.sendImgUrls=originalSend;if(AscCommon.UploadImageFiles===uploadWrapper)AscCommon.UploadImageFiles=originalUpload;}
  }};
 }
 function sendWrapper(api,images,callback,quiet,token){
  var owner=owners.get(api);
  if(owner&&owner.api===api)return owner.send(images,callback,quiet);
  return originalSend.apply(this,arguments);
 }
 function uploadWrapper(files,documentId,documentUserId,jwt,shardKey,wopiSrc,userSessionId,callback){
  var active=window.Asc&&Asc.editor||window.editor;
  var matches=Array.from(owners.values()).filter(function(candidate){
   return candidate.api===active&&candidate.api.documentId===documentId&&candidate.api.documentUserId===documentUserId&&
    candidate.api.DocInfo&&candidate.api.DocInfo.get_OfflineApp()&&candidate.api.asc_isOffline();
  });
  var owner=matches.length===1?matches[0]:null;
  if(owner)return owner.upload(Array.from(files),callback);
  return originalUpload.apply(this,arguments);
 }
 return {install:install,assertAvailable:assertAvailable};
});

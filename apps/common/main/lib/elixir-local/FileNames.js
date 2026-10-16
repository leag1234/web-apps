/*
 * SPDX-FileCopyrightText: 2026 ElixirOffice contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 */
(function(factory){
 if(typeof define==='function'&&define.amd)define([],factory);
 else if(typeof module==='object'&&module.exports)module.exports=factory();
 else {window.Common=window.Common||{};window.Common.ElixirLocalFileNames=factory();}
})(function(){
 'use strict';
 var formats=['docx','odt','rtf','txt','xlsx','ods','csv','pptx','odp','pdf'];
 function resolve(value,format){
  if(typeof value!=='string'||formats.indexOf(format)<0)return null;
  value=value.normalize('NFC');
  if(!value.trim()||/[\x00-\x1f\x7f\\/:*?"<>|\u202a-\u202e\u2066-\u2069]/.test(value)||/[. ]$/.test(value)||value.length>200)return null;
  var stem=value.trim();
  formats.some(function(extension){
   if(stem.toLowerCase().endsWith('.'+extension)){stem=stem.slice(0,-extension.length-1);return true;}
   return false;
  });
  if(!stem||stem==='.'||stem==='..'||/[. ]$/.test(stem)||/^(con|prn|aux|nul|com[1-9]|lpt[1-9])$/i.test(stem.split('.')[0].replace(/ +$/,'')))return null;
  var name=stem+'.'+format;
  return new TextEncoder().encode(name).length<=255?name:null;
 }
 return {resolve:resolve};
});

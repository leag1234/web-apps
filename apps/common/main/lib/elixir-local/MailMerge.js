/*
 * SPDX-FileCopyrightText: 2026 ElixirOffice contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 */
define([],function(){
 'use strict';
 var MAX_BYTES=20*1024*1024,MAX_RECORDS=100,MAX_COLUMNS=256,MAX_CELL=32767;
 function failure(message){var error=new Error(message);error.code='mail-merge-input';return error;}
 function value(cell){
  if(cell===null||cell===undefined)return '';
  if(typeof cell!=='string'&&typeof cell!=='number'&&typeof cell!=='boolean')throw failure('Invalid recipient cell');
  var text=String(cell);if(text.length>MAX_CELL)throw failure('Recipient cell is too long');return text;
 }
 function validate(list){
  if(!Array.isArray(list))throw failure('Recipient data needs a header row');
  function empty(row){return Array.isArray(row)&&!row.some(function(cell){return cell!==null&&cell!==undefined&&String(cell)!=='';});}
  var first=0,last=list.length-1;
  while(first<=last&&empty(list[first]))first++;
  while(last>=first&&empty(list[last]))last--;
  list=list.slice(first,last+1);
  if(!Array.isArray(list[0])||!list[0].length)throw failure('Recipient data needs a header row');
  if(list[0].length>MAX_COLUMNS)throw failure('Recipient data has too many columns');
  var rows=list.slice(0,MAX_RECORDS+1).map(function(row){
   if(!Array.isArray(row)||row.length>MAX_COLUMNS)throw failure('Invalid recipient row');
   return row.map(value);
  });
  if(!rows[0].some(function(cell){return cell.length>0;}))throw failure('Recipient data needs a header row');
  return rows;
 }
 async function recipients(input,format,converter,options){
  if(['csv','xlsx','xls'].indexOf(format)<0)throw failure('Choose a CSV or Excel recipient file');
  var bytes=input instanceof Uint8Array?input:new Uint8Array(input);
  if(!bytes.length||bytes.length>MAX_BYTES)throw failure('Choose a recipient file under 20 MiB');
  var output=await converter.convert(bytes,format,'json',options||{}),list;
  try{list=JSON.parse(new TextDecoder('utf-8',{fatal:true}).decode(output));}
  catch(_error){throw failure('Invalid recipient data');}
  return validate(list);
 }
 function afterPending(pending,run,report,alive){
  pending.then(function(){if(!alive||alive())run();}).catch(report);return true;
 }
 return {recipients:recipients,validate:validate,afterPending:afterPending,MAX_RECORDS:MAX_RECORDS};
});

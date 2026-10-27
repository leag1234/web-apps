// SPDX-License-Identifier: AGPL-3.0-only
define([], function () {
 'use strict';
 return Object.freeze({
  title:'Set up Elixir AI',
  intro:'Choose the AI service you want to use. AI stays off until you explicitly allow it.',
  disclosure:'When you ask, your request and the relevant selection or document context are sent to the service you choose. No document data is sent to the ElixirOffice publisher.',
  noChecks:'No automatic connection checks, provider discovery or telemetry.',
  provider:'AI service',
  generic:'OpenAI-compatible HTTPS service',
  endpoint:'HTTPS endpoint',
  model:'Model name',
  key:'API key',
  keyDesktop:'The key is encrypted by your operating system and is never shown again.',
  keyBrowser:'The key is stored in this browser on this computer in clear text. It is sent only to your chosen service.',
  destination:'Requests go only to the endpoint entered above.',
  continue:'Continue',
  cancel:'Cancel',
  invalid:'Enter a valid HTTPS endpoint and model name.',
  denied:'Permission was not granted. AI is still off.',
  failed:'Settings could not be saved. AI is still off.',
  waiting:'Waiting for your permission…'
 });
});

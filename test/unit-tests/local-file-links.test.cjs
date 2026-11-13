// SPDX-License-Identifier: AGPL-3.0-only
'use strict';
// Links to files of this computer in the hyperlink dialogs and on click (issue #77).
const {test} = require('node:test'), assert = require('node:assert/strict'), fs = require('node:fs'), path = require('node:path'), vm = require('node:vm');
const apps = path.resolve(__dirname, '../../apps');
const read = f => fs.readFileSync(path.join(apps, f), 'utf8');
const utils = read('common/main/lib/util/utils.js');

function loadOpener(bridge) {
	const start = utils.indexOf('utils.openLocalFileLink = function');
	const end = utils.indexOf('Common.Utils = _extend_object(Common.Utils, utils);');
	assert(start > 0 && end > start);
	const log = {alerts: [], warnings: [], opened: [], bridged: []};
	const Common = {
		UI: {alert: o => log.alerts.push(o.msg), warning: o => log.warnings.push(o)},
		Utils: {String: {format: (t, a) => t.replace('{0}', a), htmlEncode: s => s.replace(/&/g, '&amp;').replace(/</g, '&lt;')}}
	};
	const ctx = vm.createContext({Common, console, setTimeout: f => f(), window: {open: (u, t) => log.opened.push([u, t]), eoDesktop: bridge && {openLocalLink: u => log.bridged.push(u)}}});
	vm.runInContext('var utils = {};' + utils.slice(start, end) + ';this.open = utils.openLocalFileLink;', ctx);
	return {open: ctx.open, log};
}
test('the web editors say that only the desktop application opens such a link', () => {
	const {open, log} = loadOpener();
	open('file:///Users/x/doc.pdf', false, 'warn {0}');
	assert.equal(log.alerts.length, 1);
	assert.match(log.alerts[0], /desktop/);
	assert.deepEqual(log.warnings, []);
	assert.deepEqual(log.opened, []);
});
test('the desktop editor hands the link to the application, which asks natively and checks the file', () => {
	const {open, log} = loadOpener();
	open('/Users/x/a&b <doc>.pdf', true);
	assert.equal(log.warnings.length, 0, 'no page-side question: the desktop application asks itself');
	assert.deepEqual(JSON.parse(JSON.stringify(log.opened)), [['/Users/x/a&b <doc>.pdf', '_blank']]);
});
test('in the desktop application the link goes through its bridge: Chromium never opens a file: URL from the page', () => {
	const {open, log} = loadOpener(true);
	open('file:///Users/x/doc.pdf', true);
	assert.deepEqual(JSON.parse(JSON.stringify(log.bridged)), ['file:///Users/x/doc.pdf']);
	assert.deepEqual(log.opened, [], 'no window.open of a file: URL');
	assert.deepEqual(log.warnings, []);
});
test('the three hyperlink dialogs accept a link to a local file, the three editors open it on click', () => {
	for (const ed of ['spreadsheeteditor', 'documenteditor', 'presentationeditor']) {
		const dialog = read(ed + '/main/app/view/HyperlinkSettingsDialog.js');
		assert(/AscCommon\.isLocalFileHyperlink && AscCommon\.isLocalFileHyperlink\(trimmed\)\)\s*\n\s*me\.urlType = AscCommon\.c_oAscUrlType\.Unsafe/.test(dialog), ed + ' dialog');
		const holder = read(ed + '/main/app/controller/DocumentHolderExt.js');
		const own = ed === 'spreadsheeteditor' ? 'permissions' : 'mode';
		assert(new RegExp('Common\\.Utils\\.openLocalFileLink\\(url\\.trim\\(\\), this\\.' + own + ' && this\\.' + own + '\\.isElixirLocal').test(holder), ed + ' click');
		assert(holder.indexOf('openLocalFileLink') < holder.indexOf('AscCommon.getSafeHyperlink(url)'), ed + ': before the safe-link check');
		for (const l of ['en', 'ru', 'fr', 'de', 'es']) {
			const j = JSON.parse(read(ed + '/main/locale/' + l + '.json'));
			assert.equal(typeof j['Common.Utils.txtLocalLinkWeb'], 'string', ed + ' ' + l);
		}
	}
});

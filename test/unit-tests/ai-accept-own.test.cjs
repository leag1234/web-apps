// SPDX-FileCopyrightText: 2026 ElixirOffice contributors
// SPDX-License-Identifier: AGPL-3.0-only
'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');

const root = path.resolve(__dirname, '../..');
const source = fs.readFileSync(path.join(root,
    'apps/documenteditor/main/app/elixir-ai/WordTools.js'), 'utf8');

test('accept and reject use the assistant-only command when earlier revisions exist', () => {
    assert.doesNotMatch(source, /AcceptReviewChanges', \[false\]/);
    assert.doesNotMatch(source, /RejectReviewChanges', \[false\]/);
    assert.match(source, /accept: function \(ctx\) \{[\s\S]*?callCommand\(api, 'acceptOwn', \{ name: ASSISTANT_NAME, userId: 'uid-' \+ ASSISTANT_NAME \}\);/);
    assert.match(source, /reject: function \(ctx\) \{[\s\S]*?callCommand\(api, 'rejectOwn', \{ name: ASSISTANT_NAME, userId: 'uid-' \+ ASSISTANT_NAME \}\);/);
});

test('assistant revision commands collect matching author changes before applying them', () => {
    assert.match(source, /CMD\.acceptOwn = function \(a, H\) \{[\s\S]*?H\.ownRevisions\(a\);[\s\S]*?StartAction\(AscDFH\.historydescription_Document_AcceptRevisionChange, null, AscWord\.ACTION_FLAGS\.UPDATEALL_RECALCULATE\);[\s\S]*?AcceptRevisionChange\(own\[j\]\);[\s\S]*?FinalizeAction\(\);[\s\S]*?count: own\.length/);
    assert.match(source, /CMD\.rejectOwn = function \(a, H\) \{[\s\S]*?H\.ownRevisions\(a\);[\s\S]*?StartAction\(AscDFH\.historydescription_Document_RejectRevisionChange, null, AscWord\.ACTION_FLAGS\.UPDATEALL_RECALCULATE\);[\s\S]*?RejectRevisionChange\(own\[j\]\);[\s\S]*?FinalizeAction\(\);[\s\S]*?count: own\.length/);
});

test('assistant revisions are read from the per-paragraph change object (#100)', () => {
    assert.match(source, /H\.ownRevisions = function \(a\) \{[\s\S]*?Get_AllChanges\(\), own = \[\];[\s\S]*?for \(var id in byPara\)[\s\S]*?byPara\[id\]\[i\];[\s\S]*?get_UserId\(\) === a\.userId && change\.get_UserName\(\) === a\.name[\s\S]*?own\.indexOf\(change\) < 0\) own\.push\(change\);/);
});

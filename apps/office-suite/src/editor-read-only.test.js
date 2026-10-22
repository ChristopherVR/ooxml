import { test } from 'node:test';
import assert from 'node:assert/strict';
import { setEditorReadOnly } from './editor-read-only.js';

test('every product toggles readOnly and leaves canEdit alone', () => {
	for (const kind of ['docx', 'xlsx', 'vsdx']) {
		const element = { readOnly: false, canEdit: true };
		setEditorReadOnly(element, kind, true);
		assert.equal(element.readOnly, true);
		assert.equal(element.canEdit, true);
		setEditorReadOnly(element, kind, false);
		assert.equal(element.readOnly, false);
	}
});

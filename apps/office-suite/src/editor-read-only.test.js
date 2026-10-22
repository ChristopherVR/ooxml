import { test } from 'node:test';
import assert from 'node:assert/strict';
import { setEditorReadOnly } from './editor-read-only.js';

test('a Visio review turns canEdit off and restores the value it had', () => {
	const element = { canEdit: true };
	setEditorReadOnly(element, 'vsdx', true);
	assert.equal(element.canEdit, false);
	setEditorReadOnly(element, 'vsdx', true);
	assert.equal(element.canEdit, false, 'a second preview keeps the original value');
	setEditorReadOnly(element, 'vsdx', false);
	assert.equal(element.canEdit, true);
});

test('a read-only host view stays read-only when a review ends', () => {
	const element = { canEdit: false };
	setEditorReadOnly(element, 'vsdx', true);
	setEditorReadOnly(element, 'vsdx', false);
	assert.equal(element.canEdit, false);
});

test('ending a review without one starting makes the drawing editable', () => {
	const element = { canEdit: false };
	setEditorReadOnly(element, 'vsdx', false);
	assert.equal(element.canEdit, true);
});

test('Word and Excel keep their readOnly property', () => {
	for (const kind of ['docx', 'xlsx']) {
		const element = { readOnly: false, canEdit: true };
		setEditorReadOnly(element, kind, true);
		assert.equal(element.readOnly, true);
		assert.equal(element.canEdit, true);
		setEditorReadOnly(element, kind, false);
		assert.equal(element.readOnly, false);
	}
});

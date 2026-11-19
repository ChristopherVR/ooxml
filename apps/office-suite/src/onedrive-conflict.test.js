import assert from 'node:assert/strict';
import { test } from 'node:test';
import { conflictDetailsHtml } from './onedrive-conflict.js';

const escape = (text) => text.replaceAll('<', '&lt;');
const diff = (changed) => ({ changed, added: [], removed: [] });

test('lists why the edits were not combined when the packages cannot be compared', () => {
	const html = conflictDetailsHtml(
		['Both sides changed slide 2'],
		{ error: new Error('package <too large>') },
		escape,
	);
	assert.match(html, /Why the edits were not combined:/);
	assert.match(html, /Both sides changed slide 2/);
	assert.match(html, /Details are unavailable: package &lt;too large>/);
});

test('shows only the failure when there is no reason to list', () => {
	const html = conflictDetailsHtml(undefined, { error: 'unreadable' }, escape);
	assert.equal(html, '<p>Details are unavailable: unreadable</p>');
});

test('lists the reasons and the parts each side changed', () => {
	const html = conflictDetailsHtml(
		['Overlap'],
		{ remote: diff(['a.xml', 'b.xml']), local: diff(['b.xml']) },
		escape,
	);
	assert.match(html, /Overlap/);
	assert.match(html, /Changed in both:<\/p><ul><li><code>b\.xml \(changed\)/);
});

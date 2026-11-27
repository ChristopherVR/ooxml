/** The generators' regular expression escape (`samples/generators/lib/regexp.ts`). */
import assert from 'node:assert/strict';
import { test } from 'node:test';
import { escapeRegExp } from '../samples/generators/lib/regexp.ts';

test('escapeRegExp matches every syntax character literally, backslash included', () => {
	const text = String.raw`[[RAW:1]] a\b ^$.*+?(){}|/ \d`;
	for (const flags of ['', 'u']) {
		const pattern = new RegExp(`^${escapeRegExp(text)}$`, flags);
		assert.ok(pattern.test(text), flags);
		assert.ok(!pattern.test(text.replace('\\b', 'b')), flags);
	}
	assert.equal(escapeRegExp('[[RAW:12]]'), String.raw`\[\[RAW:12\]\]`);
	assert.equal(escapeRegExp('a\\'), 'a\\\\');
});

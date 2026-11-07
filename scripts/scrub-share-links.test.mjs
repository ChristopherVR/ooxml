import assert from 'node:assert/strict';
import { test } from 'node:test';

import { scrubShareLinks } from './scrub-share-links.mjs';

test('drops a session trailer and keeps the other trailers', () => {
	const input = [
		'fix(pptx): wait for webfonts',
		'',
		'Body.',
		'',
		'Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>',
		'Claude-Session: https://claude.ai/code/session_01Fri9BuoZkLqFKwZ3vcsbvk',
		'',
	].join('\n');
	const { message, removed } = scrubShareLinks(input);
	assert.deepEqual(removed, ['https://claude.ai/code/session_01Fri9BuoZkLqFKwZ3vcsbvk']);
	assert.equal(
		message,
		'fix(pptx): wait for webfonts\n\nBody.\n\nCo-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>\n',
	);
});

test('drops a line that is only a link and cuts a link out of prose', () => {
	const input =
		'docs: x\n\nSee https://chatgpt.com/share/abc-123 for the idea.\n- https://claude.ai/share/xyz\n';
	const { message, removed } = scrubShareLinks(input);
	assert.equal(removed.length, 2);
	assert.equal(message, 'docs: x\n\nSee for the idea.\n');
});

test('keeps a prose line that only looks like a trailer', () => {
	const { message } = scrubShareLinks('fix: w\n\nNote: see https://claude.ai/share/q for it.\n');
	assert.equal(message, 'fix: w\n\nNote: see for it.\n');
});

test('leaves ordinary links, comment lines and clean messages untouched', () => {
	const input =
		'fix: y\n\nRefs https://github.com/ChristopherVR/ooxml/issues/249\n# https://claude.ai/share/kept-in-comment\n';
	assert.deepEqual(scrubShareLinks(input), { message: input, removed: [] });
});

test('covers the other assistants share pages', () => {
	for (const url of [
		'https://chat.openai.com/share/1',
		'https://gemini.google.com/share/2',
		'https://g.co/gemini/share/3',
		'https://copilot.microsoft.com/shares/4',
		'https://claude.ai/chat/5',
	]) {
		assert.deepEqual(scrubShareLinks(`chore: z\n\n${url}\n`), {
			message: 'chore: z\n',
			removed: [url],
		});
	}
});

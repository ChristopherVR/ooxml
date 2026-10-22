import assert from 'node:assert/strict';
import { test } from 'node:test';

import { buildPayload, cleanNotes, releasedPackages, truncate } from './announce-discord.mjs';

const release = (key, npm, version, notes = '### Features\n\n- add a thing (abc1234)') => ({
	key,
	npm,
	version,
	url: `https://github.com/o/r/releases/tag/${npm}@${version}`,
	notes,
});

test('truncate keeps short text and cuts long text on a line boundary', () => {
	assert.equal(truncate('short', 20), 'short');
	const out = truncate('line one\nline two\nline three', 20);
	assert.ok(out.length <= 20);
	assert.ok(out.endsWith('…'));
	assert.ok(!out.includes('line thr'));
});

test('cleanNotes drops the pruned-history footer', () => {
	assert.equal(
		cleanNotes('### Changes\n\n- a\n\n---\nSuperseded releases are pruned'),
		'### Changes\n\n- a',
	);
	assert.equal(cleanNotes(undefined), '');
});

test('releasedPackages keeps only released packages, in plan order', () => {
	const plan = {
		order: ['core', 'ui', 'mcp'],
		packages: {
			core: { npm: 'ooxml-core', release: true },
			ui: { npm: 'ooxml-ui', release: false },
			mcp: { npm: 'ooxml-mcp', release: true },
		},
	};
	assert.deepEqual(
		releasedPackages(plan).map((p) => p.key),
		['core', 'mcp'],
	);
});

test('a successful run lists every package and quotes the highlighted notes', () => {
	const payload = buildPayload({
		repo: 'o/r',
		status: 'success',
		releases: [
			release('core', 'ooxml-core', '1.2.0'),
			release('react', 'pptx-react-viewer', '4.45.0'),
		],
	});
	assert.deepEqual(payload.allowed_mentions, { parse: [] });
	assert.equal(payload.embeds[0].title, 'Released 2 packages');
	assert.match(payload.embeds[0].description, /pptx-react-viewer\*\* \[4\.45\.0\]/u);
	assert.equal(payload.embeds.length, 2);
	assert.equal(payload.embeds[1].title, 'ooxml-core 1.2.0');
});

test('a failed publish says so and links the run', () => {
	const payload = buildPayload({
		repo: 'o/r',
		status: 'failure',
		runUrl: 'https://github.com/o/r/actions/runs/1',
		releases: [release('core', 'ooxml-core', '1.2.0')],
	});
	assert.match(payload.embeds[0].title, /did not finish/u);
	assert.match(payload.embeds[0].description, /actions\/runs\/1/u);
});

test('the message stays within Discord limits however many packages and notes there are', () => {
	const big = '- a very long changelog line that keeps going\n'.repeat(200);
	const many = Array.from({ length: 60 }, (_, i) => release(`v${i}`, `pkg-viewer-${i}`, '1.0.0'));
	const payload = buildPayload({
		repo: 'o/r',
		status: 'success',
		releases: [
			release('core', 'ooxml-core', '1.0.0', big),
			release('ui', 'ooxml-ui', '1.0.0', big),
			release('mcp', 'ooxml-mcp', '1.0.0', big),
			...many,
		],
	});
	const total = payload.embeds.reduce(
		(sum, e) => sum + e.title.length + e.description.length + (e.footer?.text.length ?? 0),
		0,
	);
	assert.ok(total <= 6000, `total ${total}`);
	assert.ok(payload.embeds.length <= 10);
	for (const e of payload.embeds) assert.ok(e.description.length <= 4096);
});

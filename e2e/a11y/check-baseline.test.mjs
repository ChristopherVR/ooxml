import assert from 'node:assert/strict';
import { spawnSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import { test } from 'node:test';
import {
	grown,
	grownNotices,
	loosenedGaps,
	lowered,
	openGaps,
	unmappedNotices,
	unowned,
} from './check-baseline.mjs';

const base = { allowed: { 'xlsx/light': { 'aria-required-parent': 47, 'target-size': 6 } } };

test('a baseline that allows more than its base is refused', () => {
	assert.deepEqual(grown(base, base), []);
	assert.deepEqual(grown(base, { allowed: { 'xlsx/light': { 'aria-required-parent': 40 } } }), []);
	assert.deepEqual(grown(base, { allowed: { 'xlsx/light': { 'aria-required-parent': 48 } } }), [
		'xlsx/light: aria-required-parent rose from 47 to 48 nodes',
	]);
	assert.deepEqual(grown(base, { allowed: { 'docx/dark': { 'color-contrast': 1 } } }), [
		'docx/dark: color-contrast is new (1 nodes)',
	]);
});

test('a known gap may not be added, left vague or changed to expect another failure', () => {
	const before = { gaps: { 'keyboard/vsdx/f6': { fails: 'F6 never came back' } } };
	assert.deepEqual(loosenedGaps(before, before), []);
	assert.deepEqual(loosenedGaps(before, { gaps: {} }), []);
	assert.deepEqual(
		loosenedGaps(before, {
			gaps: { ...before.gaps, 'keyboard/docx/tab': { fails: 'Tab stuck' } },
		}),
		['keyboard/docx/tab is new'],
	);
	assert.deepEqual(loosenedGaps(before, { gaps: { 'keyboard/vsdx/f6': { fails: 'F6' } } }), [
		'keyboard/vsdx/f6 expects a different failure',
	]);
	assert.deepEqual(loosenedGaps(before, { gaps: { 'keyboard/vsdx/f6': {} } }), [
		'keyboard/vsdx/f6 does not say which failure it expects',
	]);
	// A base entry without `fails` matched any failure, so naming one only narrows it.
	assert.deepEqual(loosenedGaps({ gaps: { 'keyboard/vsdx/f6': {} } }, before), []);
});

test('a sample notice must quote an open item of the parity checklist', () => {
	const checklist = [
		'## Word',
		'- [ ] **P1** A chart in a document cannot be edited (data, type or',
		'      formatting), and the load note still calls it a placeholder.',
		'- [x] **P2** Text boxes are drawn on pages.',
		'- [ ] **P2** Arrowhead styles outside the common set are not drawn.',
	].join('\n');
	assert.equal(openGaps(checklist).length, 2);
	const entry = (notice, gap) => ({ notices: { sample: [{ notice, gap }] } });
	// The quote may run across a line break of the list.
	assert.deepEqual(
		unmappedNotices(entry('Chart', 'cannot be edited (data, type or formatting)'), checklist),
		[],
	);
	assert.deepEqual(unmappedNotices(entry('Text box', 'Text boxes are drawn on pages'), checklist), [
		'sample: "Text box" quotes no open item of docs/parity-checklist.md',
	]);
	assert.deepEqual(unmappedNotices(entry('Chart'), checklist), [
		'sample: "Chart" quotes no open item of docs/parity-checklist.md',
	]);
	assert.deepEqual(unmappedNotices({ notices: { sample: [{ gap: 'Arrowhead' }] } }, checklist), [
		'sample: an entry has no notice',
	]);
});

test('the notices a sample may show, and the standing notes ignored, may only shrink', () => {
	const before = {
		informational: ['Fields show'],
		notices: {
			report: [
				{ notice: 'Chart', gap: 'x' },
				{ notice: 'Merged', gap: 'y' },
			],
		},
	};
	assert.deepEqual(grownNotices(before, before), []);
	assert.deepEqual(
		grownNotices(before, { informational: [], notices: { report: [before.notices.report[0]] } }),
		[],
	);
	assert.deepEqual(
		grownNotices(before, {
			informational: ['Fields show', 'Theme colors'],
			notices: { report: [...before.notices.report, { notice: 'Text box', gap: 'z' }] },
		}),
		['informational "Theme colors" is new', 'report: "Text box" is new'],
	);
	// A sample the base did not have brings its own notices.
	assert.deepEqual(
		grownNotices(before, {
			notices: { ...before.notices, form: [{ notice: 'Legacy', gap: 'w' }] },
		}),
		[],
	);
});

test('a base git cannot find fails the check instead of passing it', () => {
	const run = spawnSync(
		process.execPath,
		[fileURLToPath(new URL('check-baseline.mjs', import.meta.url)), '--base', 'no-such-ref/x'],
		{ encoding: 'utf8' },
	);
	assert.equal(run.status, 1);
	assert.match(run.stderr, /git cannot find the base no-such-ref\/x/u);
});

test('updating lowers counts and drops fixed rules, and raises nothing without --add', () => {
	const observed = { 'xlsx/light': { 'aria-required-parent': 50, 'color-contrast': 2 } };
	assert.deepEqual(lowered(base, observed).allowed, {
		'xlsx/light': { 'aria-required-parent': 47 },
	});
	assert.deepEqual(lowered(base, { 'xlsx/light': {} }).allowed, {});
	assert.deepEqual(lowered(base, { 'docx/light': { 'target-size': 3 } }).allowed, base.allowed);
	assert.deepEqual(lowered(base, observed, true).allowed, {
		'xlsx/light': { 'aria-required-parent': 47, 'color-contrast': 2 },
	});
});

test('every allowed rule names the package that removes it', () => {
	assert.deepEqual(unowned({ ...base, owners: { 'aria-required-parent': 'P4' } }), ['target-size']);
});

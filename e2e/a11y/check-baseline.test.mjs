import assert from 'node:assert/strict';
import { spawnSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import { test } from 'node:test';
import { grown, loosenedGaps, lowered, unowned } from './check-baseline.mjs';

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

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { addedGaps, grown, lowered, unowned } from './check-baseline.mjs';

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

test('a known gap the base did not list is refused', () => {
	const before = { gaps: { 'keyboard/vsdx/f6': {} } };
	assert.deepEqual(addedGaps(before, before), []);
	assert.deepEqual(addedGaps(before, { gaps: {} }), []);
	assert.deepEqual(
		addedGaps(before, { gaps: { 'keyboard/vsdx/f6': {}, 'keyboard/docx/tab': {} } }),
		['keyboard/docx/tab'],
	);
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

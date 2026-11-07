import assert from 'node:assert/strict';
import { test } from 'node:test';

import {
	DEFAULT_SETTINGS,
	authorName,
	editorLocale,
	editorRadius,
	externalContentOption,
	normalizeSettings,
	onSettings,
	resetSettings,
	settings,
	updateSettings,
} from './settings.js';
import { settingsCategories, settingsDefaults } from './settings-model.js';

test('the defaults leave every editor on its own choice', () => {
	assert.deepEqual(settings(), DEFAULT_SETTINGS);
	assert.equal(editorLocale(), undefined);
	assert.equal(externalContentOption(), undefined);
	assert.equal(editorRadius(), '4px');
});

test('saved values are checked: wrong types, unknown choices and keys fall back', () => {
	assert.deepEqual(
		normalizeSettings({
			locale: 'xx',
			corners: 'blob',
			showRibbon: 'no',
			measurementUnit: 'centimetres',
			authorName: '  Ada\nLovelace  ',
			extra: 1,
		}),
		{ ...DEFAULT_SETTINGS, measurementUnit: 'centimetres', authorName: 'Ada Lovelace' },
	);
	assert.equal(normalizeSettings({ authorName: 'x'.repeat(200) }).authorName.length, 80);
	assert.deepEqual(normalizeSettings('nonsense'), DEFAULT_SETTINGS);
});

test('updates notify with the changed keys only, and reset restores the defaults', () => {
	const seen = [];
	const stop = onSettings((_next, keys) => seen.push(keys));
	updateSettings({ locale: 'de', showRibbon: true });
	updateSettings({ locale: 'de' });
	assert.deepEqual(seen, [['locale']]);
	assert.equal(editorLocale(), 'de');
	updateSettings({ blockExternalContent: true, corners: 'square' });
	assert.equal(externalContentOption(), false);
	assert.equal(editorRadius(), '0px');
	resetSettings();
	assert.deepEqual(settings(), DEFAULT_SETTINGS);
	stop();
});

test('the author is the setting, else the profile name, else nothing', () => {
	assert.equal(authorName('Grace', { ...DEFAULT_SETTINGS, authorName: 'Ada' }), 'Ada');
	assert.equal(authorName(' Grace ', DEFAULT_SETTINGS), 'Grace');
	assert.equal(authorName('', DEFAULT_SETTINGS), undefined);
});

test('every row of the dialog edits a real setting, and Assistant shows only when enabled', () => {
	const categories = settingsCategories({ profileName: 'Ada' });
	const keys = categories.flatMap((c) => c.sections.flatMap((s) => s.controls.map((x) => x.key)));
	assert.deepEqual(new Set(keys), new Set(Object.keys(settingsDefaults())));
	assert.equal(new Set(keys).size, keys.length);
	assert.deepEqual(
		categories.map((c) => c.id),
		['general', 'view', 'language', 'fonts', 'privacy'],
	);
	assert.match(JSON.stringify(categories), /profile name \(Ada\)/);
	assert.equal(settingsCategories({ assistant: true }).at(-1).id, 'assistant');
});

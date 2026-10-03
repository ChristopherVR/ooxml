import assert from 'node:assert/strict';
import { test } from 'node:test';
import { APPS, demoUrl, isLive } from '../site/apps.js';
import { currentTheme, initTheme } from '../site/theme.js';

test('Excel beta launches the six deployed viewer routes', () => {
	const excel = APPS.find((app) => app.id === 'excel');
	assert(excel);
	assert(isLive(excel));
	assert.equal(excel.tag.label, 'Beta');
	assert.equal(excel.docs, 'https://christophervr.github.io/xlsx-viewer/');
	assert.deepEqual(
		Object.fromEntries(
			excel.frameworks.map((framework) => [framework.id, demoUrl(excel, framework)]),
		),
		{
			react: 'https://christophervr.github.io/xlsx-viewer/demo/',
			vue: 'https://christophervr.github.io/xlsx-viewer/demo-vue/',
			angular: 'https://christophervr.github.io/xlsx-viewer/demo-angular/',
			svelte: 'https://christophervr.github.io/xlsx-viewer/demo-svelte/',
			vanilla: 'https://christophervr.github.io/xlsx-viewer/demo-vanilla/',
			solid: 'https://christophervr.github.io/xlsx-viewer/demo-solid/',
		},
	);
});

test('suite appearance writes the Excel shared preference without reloading its workbook', (t) => {
	const handlers = new Map();
	const store = new Map();
	const original = new Map(
		['document', 'window', 'localStorage', 'matchMedia'].map((key) => [key, globalThis[key]]),
	);
	t.after(() => {
		for (const [key, value] of original) {
			if (value === undefined) delete globalThis[key];
			else globalThis[key] = value;
		}
	});
	globalThis.document = {
		documentElement: { dataset: { theme: 'light' } },
		getElementById: () => ({ addEventListener: (event, handler) => handlers.set(event, handler) }),
	};
	globalThis.window = { addEventListener: (event, handler) => handlers.set(event, handler) };
	globalThis.localStorage = {
		getItem: (key) => store.get(key) ?? null,
		setItem: (key, value) => store.set(key, value),
	};
	globalThis.matchMedia = () => ({ matches: false });
	let reloads = 0;
	const frame = {
		contentDocument: { querySelector: () => null },
		contentWindow: { location: { reload: () => reloads++ } },
	};
	initTheme(frame);
	handlers.get('click')();
	assert.equal(currentTheme(), 'dark');
	assert.equal(store.get('vitepress-theme-appearance'), 'dark');
	assert.equal(reloads, 0);
	handlers.get('storage')({ key: 'vitepress-theme-appearance', newValue: 'light' });
	assert.equal(currentTheme(), 'light');
	assert.equal(reloads, 0);
});

import { test } from 'node:test';
import assert from 'node:assert/strict';
import { JSDOM } from 'jsdom';
import { initVisioTheme } from './assets/theme.js';

function setup(saved, systemDark = false) {
	const dom = new JSDOM('<button class="theme-toggle"></button>', { url: 'https://example.test' });
	const media = new dom.window.EventTarget();
	media.matches = systemDark;
	dom.window.matchMedia = () => media;
	if (saved) dom.window.localStorage.setItem('vitepress-theme-appearance', saved);
	const dispose = initVisioTheme(dom.window);
	return { dom, media, dispose, root: dom.window.document.documentElement };
}

test('shell preference overrides the legacy Visio preference and updates live', () => {
	const { dom, root, dispose } = setup('dark');
	assert.equal(root.dataset.theme, 'dark');
	dom.window.dispatchEvent(
		new dom.window.StorageEvent('storage', {
			key: 'vitepress-theme-appearance',
			newValue: 'light',
		}),
	);
	assert.equal(root.dataset.theme, 'light');
	assert.equal(
		dom.window.document.querySelector('button').getAttribute('aria-label'),
		'Switch to dark theme',
	);
	dispose();
	dom.window.dispatchEvent(
		new dom.window.StorageEvent('storage', {
			key: 'vitepress-theme-appearance',
			newValue: 'dark',
		}),
	);
	assert.equal(root.dataset.theme, 'light');
	dom.window.close();
});

test('auto and cleared storage follow system changes; explicit choices take priority', () => {
	const { dom, root, media } = setup('auto');
	assert.equal(root.dataset.theme, 'light');
	media.matches = true;
	media.dispatchEvent(new dom.window.Event('change'));
	assert.equal(root.style.colorScheme, 'dark');
	dom.window.document.querySelector('button').click();
	assert.equal(root.dataset.theme, 'light');
	media.dispatchEvent(new dom.window.Event('change'));
	assert.equal(root.dataset.theme, 'light');
	dom.window.dispatchEvent(new dom.window.StorageEvent('storage', { key: null }));
	assert.equal(root.dataset.theme, 'dark');
	assert.equal(root.style.colorScheme, 'dark');
	dom.window.close();
});

test('theme toggle works when preference storage is unavailable', () => {
	const { dom, root, dispose } = setup('light');
	dispose();
	Object.defineProperty(dom.window, 'localStorage', {
		get() {
			throw new Error('blocked');
		},
	});
	initVisioTheme(dom.window);
	dom.window.document.querySelector('button').click();
	assert.equal(root.dataset.theme, 'dark');
	dom.window.close();
});

import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { wireWorkspaceTheme } from '../demo/workspace-theme.js';

let dispose: (() => void) | undefined;
beforeEach(() => {
	vi.stubGlobal('matchMedia', () => ({
		matches: false,
		addEventListener() {},
		removeEventListener() {},
	}));
});
afterEach(() => {
	dispose?.();
	dispose = undefined;
	localStorage.clear();
	document.body.replaceChildren();
	delete document.documentElement.dataset.theme;
	delete document.documentElement.dataset.embedded;
	history.replaceState(null, '', '/');
	vi.unstubAllGlobals();
});
function mount(): HTMLButtonElement {
	document.documentElement.dataset.theme = 'dark';
	document.body.innerHTML = '<button id="theme-toggle" type="button">Theme</button>';
	dispose = wireWorkspaceTheme(document);
	return document.querySelector('button')!;
}
describe('workspace appearance', () => {
	it('shares the documentation preference and updates accessible labels', () => {
		localStorage.setItem('visio-docs-theme', 'light');
		const button = mount();
		expect(document.documentElement.dataset.theme).toBe('light');
		expect(button.getAttribute('aria-label')).toBe('Switch to dark theme');
		button.click();
		expect(document.documentElement.dataset.theme).toBe('dark');
		expect(localStorage.getItem('visio-docs-theme')).toBe('dark');
		expect(button.getAttribute('title')).toBe('Switch to light theme');
	});
	it('receives same-origin appearance changes and removes listeners on disposal', () => {
		const button = mount();
		window.dispatchEvent(
			new StorageEvent('storage', { key: 'visio-docs-theme', newValue: 'light' }),
		);
		expect(document.documentElement.dataset.theme).toBe('light');
		dispose?.();
		button.click();
		expect(document.documentElement.dataset.theme).toBe('light');
		window.dispatchEvent(
			new StorageEvent('storage', { key: 'visio-docs-theme', newValue: 'dark' }),
		);
		expect(document.documentElement.dataset.theme).toBe('light');
	});
	it('supports embedded compact chrome without changing document contents', () => {
		history.replaceState(null, '', '/demo/?embed=1');
		mount();
		expect(document.documentElement.hasAttribute('data-embedded')).toBe(true);
		expect(document.querySelectorAll('button')).toHaveLength(1);
	});
});

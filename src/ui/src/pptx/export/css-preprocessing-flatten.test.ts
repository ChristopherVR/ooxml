// @vitest-environment jsdom
/**
 * Regression tests for `removeUnsupportedFeatures`: html2canvas-pro 2.6 rejects a clone whose
 * computed `contain` or `will-change` holds several values ("multiple values found when expecting
 * only one"), so the cleanup pass keeps one value on each cloned element.
 */
import { afterEach, describe, expect, it, vi } from 'vitest';

import { removeUnsupportedFeatures } from './css-preprocessing-flatten';

afterEach(() => {
	vi.restoreAllMocks();
	document.body.innerHTML = '';
});

/** Make the computed style report `values` for the named properties, empty for the rest. */
function stubComputed(values: Record<string, string>): void {
	const real = window.getComputedStyle.bind(window);
	vi.spyOn(window, 'getComputedStyle').mockImplementation((el: Element) => {
		const base = real(el);
		return {
			getPropertyValue: (name: string) => values[name] ?? base.getPropertyValue(name),
		} as unknown as CSSStyleDeclaration;
	});
}

function clonedWith(tag = 'div'): HTMLElement {
	const root = document.createElement('section');
	const child = document.createElement(tag);
	root.appendChild(child);
	document.body.appendChild(root);
	return root;
}

describe('removeUnsupportedFeatures: multi-value rendering hints', () => {
	it('keeps the paint keyword of a multi-keyword contain', () => {
		stubComputed({ contain: 'layout paint' });
		const root = clonedWith();
		removeUnsupportedFeatures(root);
		const child = root.firstElementChild as HTMLElement;
		expect(child.style.getPropertyValue('contain')).toBe('paint');
	});

	it('keeps the first keyword when a multi-keyword contain has no paint', () => {
		stubComputed({ contain: 'layout style' });
		const root = clonedWith();
		removeUnsupportedFeatures(root);
		const child = root.firstElementChild as HTMLElement;
		expect(child.style.getPropertyValue('contain')).toBe('layout');
	});

	it('resets a comma-separated will-change to auto', () => {
		stubComputed({ 'will-change': 'transform, opacity' });
		const root = clonedWith();
		removeUnsupportedFeatures(root);
		const child = root.firstElementChild as HTMLElement;
		expect(child.style.getPropertyValue('will-change')).toBe('auto');
	});

	it('leaves a single-value will-change and an absent contain alone', () => {
		stubComputed({ 'will-change': 'transform', contain: 'none' });
		const root = clonedWith();
		removeUnsupportedFeatures(root);
		const child = root.firstElementChild as HTMLElement;
		expect(child.style.getPropertyValue('will-change')).toBe('');
		expect(child.style.getPropertyValue('contain')).toBe('');
	});
});

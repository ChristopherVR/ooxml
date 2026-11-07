// @vitest-environment happy-dom
import type { TextSegment, TextStyle } from 'ooxml-core/pptx';
import { afterEach, describe, expect, it } from 'vitest';

import { seedPlainInlineText } from './inline-text-seed';

afterEach(() => {
	document.body.replaceChildren();
});

function surface(): HTMLElement {
	const el = document.createElement('div');
	document.body.append(el);
	return el;
}

describe('seedPlainInlineText', () => {
	it('gives each run the substituted font chain, never the bare deck name', () => {
		const segments: TextSegment[] = [
			{ text: 'The Year Is', style: { fontFamily: 'Playfair Display', fontSize: 62 } },
		];
		const root = seedPlainInlineText(surface(), segments, 'The Year Is');
		const span = root.querySelector<HTMLElement>('[data-seg-idx="0"]');
		expect(span?.textContent).toBe('The Year Is');
		const family = span?.style.fontFamily ?? '';
		expect(family.startsWith('"Playfair Display"')).toBe(true);
		// The chain ends in a generic family, so a deck font that is not painted
		// falls back exactly as it does in view mode.
		expect(family).toMatch(/(sans-serif|serif|monospace)$/u);
	});

	it('applies the body autofit scale to run sizes', () => {
		const segments: TextSegment[] = [{ text: 'Shrunk', style: { fontSize: 40 } }];
		const textStyle: TextStyle = {
			fontSize: 40,
			autoFit: true,
			autoFitMode: 'normal',
			autoFitFontScale: 0.5,
		};
		const root = seedPlainInlineText(surface(), segments, 'Shrunk', textStyle);
		const span = root.querySelector<HTMLElement>('[data-seg-idx="0"]');
		expect(Number.parseFloat(span?.style.fontSize ?? '')).toBeLessThan(40);
	});

	it('keeps one text-flow container and the plain-text fallback', () => {
		const withSegments = surface();
		const root = seedPlainInlineText(
			withSegments,
			[
				{ text: 'a', style: {} },
				{ text: 'b', style: {} },
			],
			'ab',
		);
		expect(root.dataset.pptxTextFlow).toBe('');
		expect(withSegments.children).toHaveLength(1);
		expect(root.textContent).toBe('ab');

		const plain = surface();
		expect(seedPlainInlineText(plain, undefined, 'plain text')).toBe(plain);
		expect(plain.textContent).toBe('plain text');
	});
});

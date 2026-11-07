import { readFileSync } from 'node:fs';

import { describe, expect, it } from 'vitest';

import type { PptxEmbeddedFont, PptxSlide } from '../../types';
import { PptxHandler } from '../../PptxHandler';

function deck(path: string): ArrayBuffer {
	const bytes = readFileSync(new URL(`../../../__tests__/fixtures/${path}`, import.meta.url));
	return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
}

describe('PptxHandler.load onSlidesParsed', () => {
	it('reports the parsed slides and embedded fonts once, before the load resolves', async () => {
		const handler = new PptxHandler();
		const calls: { slides: readonly PptxSlide[]; embeddedFonts: readonly PptxEmbeddedFont[] }[] =
			[];
		let resolvedBeforeHook = false;
		const pending = handler.load(deck('e2e/adlam-webfont.pptx'), {
			onSlidesParsed: (parsed) => {
				calls.push(parsed);
			},
		});
		void pending.then(() => {
			resolvedBeforeHook = calls.length === 0;
		});
		const data = await pending;
		try {
			expect(resolvedBeforeHook).toBe(false);
			expect(calls).toHaveLength(1);
			const [parsed] = calls;
			expect(parsed?.slides.map((slide) => slide.id)).toEqual(data.slides.map((slide) => slide.id));
			expect(parsed?.embeddedFonts).toEqual(data.embeddedFonts ?? []);
			// The slides already carry the resolved font a webfont download needs.
			const families = JSON.stringify(parsed?.slides);
			expect(families).toContain('ADLaM Display');
		} finally {
			handler.dispose();
		}
	});

	it('loads exactly as before without the hook', async () => {
		const handler = new PptxHandler();
		const data = await handler.load(deck('e2e/adlam-webfont.pptx'));
		try {
			expect(data.slides.length).toBeGreaterThan(0);
		} finally {
			handler.dispose();
		}
	});
});

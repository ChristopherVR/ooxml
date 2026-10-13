import JSZip from 'jszip';
import { expect, it } from 'vitest';
import { PptxHandler } from '../../core/PptxHandler';
import { createVisibilityTimingDeck } from '../helpers/visibility-timing-deck';

it.each([0, 499, 1499])(
	'preserves native timing at %ims through no-op and unrelated text edits',
	async (delay) => {
		const source = await createVisibilityTimingDeck(delay);
		const original = await JSZip.loadAsync(source);
		const xml = await original.file('ppt/slides/slide1.xml')!.async('string');
		const timing = xml.match(/<p:timing\b[\s\S]*?<\/p:timing>/u)?.[0];
		const handler = new PptxHandler();
		try {
			const data = await handler.load(
				source.buffer.slice(source.byteOffset, source.byteOffset + source.byteLength),
			);
			const slide = data.slides[0]!;
			expect(slide.nativeAnimations?.[0]).toMatchObject({
				durationMs: delay + 1,
				visibilitySetTiming: { delayMs: delay, durationMs: 1 },
			});
			for (const edit of [false, true]) {
				const shape = slide.elements.find(
					(element) => element.type === 'shape' && element.text === 'RectDimRGB',
				);
				if (!shape || shape.type !== 'shape') throw new Error('Missing native target');
				if (edit) {
					shape.text = 'Edited native target';
					shape.textSegments = [{ text: shape.text, style: {} }];
				}
				const output = await handler.save(data.slides);
				const saved = await JSZip.loadAsync(output);
				expect(
					(await saved.file('ppt/slides/slide1.xml')!.async('string')).match(
						/<p:timing\b[\s\S]*?<\/p:timing>/u,
					)?.[0],
				).toBe(timing);
				const reopened = new PptxHandler();
				try {
					expect((await reopened.load(output)).slides[0]?.nativeAnimations?.[0]?.durationMs).toBe(
						delay + 1,
					);
				} finally {
					reopened.dispose();
				}
			}
		} finally {
			handler.dispose();
		}
	},
);

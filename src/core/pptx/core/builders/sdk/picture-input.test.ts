/**
 * A new picture's opacity and brightness / contrast reach the file as
 * `a:alphaModFix` and `a:lum` and load back as the same picture effects.
 */
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';

import { PptxHandler } from '../../PptxHandler';
import type { PptxImageEffects } from '../../types/image';
import { createImageElement } from './ElementFactory';
import type { ImageOptions } from './types';

const PNG =
	'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';

async function roundTrip(
	options: ImageOptions,
): Promise<{ blip: string; effects: PptxImageEffects | undefined; opacity: number | undefined }> {
	const { handler, createSlide } = await PptxHandler.create();
	const bytes = await handler.save([createSlide('Blank').addImage(PNG, options).build()]);
	const zip = await JSZip.loadAsync(bytes);
	const xml = (await zip.file('ppt/slides/slide1.xml')?.async('string')) ?? '';
	const loaded = await new PptxHandler().load(bytes);
	const pic = loaded.slides[0]?.elements[0] as
		| { imageEffects?: PptxImageEffects; opacity?: number }
		| undefined;
	return {
		blip: /<a:blip\b[\s\S]*?(?:\/>|<\/a:blip>)/u.exec(xml)?.[0] ?? '',
		effects: pic?.imageEffects,
		opacity: pic?.opacity,
	};
}

describe('picture opacity', () => {
	it('writes opacity as a:alphaModFix and reads it back', async () => {
		const { blip, effects } = await roundTrip({ opacity: 0.4 });
		expect(blip).toContain('<a:alphaModFix amt="40000"');
		expect(effects?.alphaModFix).toBe(40);
	});

	it('writes nothing for full or omitted opacity', async () => {
		for (const options of [{ opacity: 1 }, {}]) {
			const { blip, effects } = await roundTrip(options);
			expect(blip).not.toContain('alphaModFix');
			expect(effects?.alphaModFix).toBeUndefined();
		}
	});

	it('writes a fully transparent picture as amt 0', async () => {
		const { blip, effects } = await roundTrip({ opacity: 0 });
		expect(blip).toContain('<a:alphaModFix amt="0"');
		expect(effects?.alphaModFix).toBe(0);
	});

	it('keeps a fractional opacity to the file precision', async () => {
		const { blip, effects } = await roundTrip({ opacity: 0.123456 });
		expect(blip).toContain('<a:alphaModFix amt="12346"');
		expect(effects?.alphaModFix).toBeCloseTo(12.346, 3);
	});

	it('applies the opacity once: as the picture effect, not also on the element', () => {
		const el = createImageElement(PNG, { opacity: 0.5 });
		expect(el.opacity).toBeUndefined();
		expect(el.imageEffects?.alphaModFix).toBe(50);
	});

	it('rejects an opacity outside 0..1', () => {
		for (const opacity of [-0.1, 1.5, Number.NaN]) {
			expect(() => createImageElement(PNG, { opacity })).toThrow(RangeError);
		}
	});
});

describe('picture brightness and contrast', () => {
	it('writes them as a:lum, not the legacy blip attributes, and reads them back', async () => {
		const { blip, effects } = await roundTrip({ brightness: 20, contrast: -10 });
		expect(blip).toContain('<a:lum bright="20000" contrast="-10000"');
		expect(blip).not.toMatch(/<a:blip [^>]*bright=/u);
		expect(effects?.lum).toEqual({ bright: 20, contrast: -10 });
	});

	it('writes only the value that is given, including the extremes', async () => {
		const bright = await roundTrip({ brightness: 100 });
		expect(bright.blip).toContain('<a:lum bright="100000"');
		expect(bright.blip).not.toContain('contrast=');
		const contrast = await roundTrip({ contrast: -100 });
		expect(contrast.blip).toMatch(/<a:lum contrast="-100000"/u);
		expect(contrast.effects?.lum?.bright).toBeUndefined();
	});

	it('writes a zero change as given', async () => {
		const { blip } = await roundTrip({ brightness: 0, contrast: 0 });
		expect(blip).toContain('<a:lum bright="0" contrast="0"');
	});

	it('writes opacity and luminance together', async () => {
		const { effects } = await roundTrip({ opacity: 0.75, brightness: 12.5 });
		expect(effects?.alphaModFix).toBe(75);
		expect(effects?.lum?.bright).toBe(12.5);
	});

	it('rejects a change outside -100..100', () => {
		expect(() => createImageElement(PNG, { brightness: 101 })).toThrow(RangeError);
		expect(() => createImageElement(PNG, { contrast: -150 })).toThrow(RangeError);
		expect(() => createImageElement(PNG, { contrast: Number.POSITIVE_INFINITY })).toThrow(
			RangeError,
		);
	});
});

import type { PptxElement } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import { TEXT_WARP_PRESETS } from '../text-warp-presets';
import { applyRibbonGalleryItem, buildRibbonGallery } from './gallery-registry';
import { textWarpStyle } from './text-warp-gallery';

const text = (textWarpPreset?: string): PptxElement =>
	({
		id: 'title',
		type: 'text',
		x: 0,
		y: 0,
		width: 200,
		height: 60,
		text: 'Hello',
		textStyle: { fontSize: 24, ...(textWarpPreset ? { textWarpPreset } : {}) },
	}) as PptxElement;

describe('text warp gallery', () => {
	it('offers every preset as a tile and marks the current one', () => {
		const gallery = buildRibbonGallery('textWarp', { element: text('textArchUp') });
		const items = gallery.sections.flatMap((section) => section.items);
		expect(items.map((item) => item.id)).toStrictEqual(TEXT_WARP_PRESETS.map((p) => p.value));
		expect(items.filter((item) => item.applied).map((item) => item.id)).toStrictEqual([
			'textArchUp',
		]);
		expect(items[0]!.previewSvg).toMatch(/^<svg [^>]*viewBox="0 0 40 20"/);
		expect(gallery.disabled).toBeFalsy();
	});

	it('marks No Transform when the text is not warped, and is unavailable without text', () => {
		const items = buildRibbonGallery('textWarp', { element: text() }).sections[0]!.items;
		expect(items.find((item) => item.applied)?.id).toBe('textNoShape');
		const picture = { id: 'pic', type: 'image', x: 0, y: 0, width: 1, height: 1 } as PptxElement;
		expect(buildRibbonGallery('textWarp', { element: picture }).disabled).toBeTruthy();
		expect(buildRibbonGallery('textWarp', { element: null }).disabled).toBeTruthy();
		expect(applyRibbonGalleryItem('textWarp', 'textArchUp', { element: picture })).toBeNull();
	});

	it('sets the preset on the text style, and No Transform clears it', () => {
		expect(applyRibbonGalleryItem('textWarp', 'textWave1', { element: text() })).toStrictEqual({
			kind: 'element',
			elementId: 'title',
			patch: { textStyle: { fontSize: 24, textWarpPreset: 'textWave1' } },
		});
		expect(
			applyRibbonGalleryItem('textWarp', 'textNoShape', { element: text('textWave1') }),
		).toStrictEqual({
			kind: 'element',
			elementId: 'title',
			patch: { textStyle: { fontSize: 24, textWarpPreset: undefined } },
		});
		expect(applyRibbonGalleryItem('textWarp', 'archUp', { element: text() })).toBeNull();
	});

	it('names the text style change of a tile for hosts that update the text style', () => {
		expect(textWarpStyle('textArchUp')).toStrictEqual({ textWarpPreset: 'textArchUp' });
		expect(textWarpStyle('textNoShape')).toStrictEqual({ textWarpPreset: undefined });
		// Short names are not OOXML presets: the renderer would not know them.
		expect(textWarpStyle('archUp')).toBeUndefined();
	});
});

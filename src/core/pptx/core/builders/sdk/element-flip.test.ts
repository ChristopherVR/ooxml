/**
 * Flips on new elements are saved as `flipH` / `flipV` (not folded into a
 * rotation) and load back as the same flips.
 */
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';

import { PptxHandler } from '../../PptxHandler';
import type { PptxElement } from '../../types';
import { createMediaElement, createShapeElement } from './ElementFactory';
import type { SlideBuilder } from './SlideBuilder';

const PNG =
	'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';

async function roundTrip(
	build: (slide: SlideBuilder) => void,
): Promise<{ xfrms: string[]; elements: PptxElement[] }> {
	const { handler, createSlide } = await PptxHandler.create();
	const slide = createSlide('Blank');
	build(slide);
	const bytes = await handler.save([slide.build()]);
	const zip = await JSZip.loadAsync(bytes);
	const xml = (await zip.file('ppt/slides/slide1.xml')?.async('string')) ?? '';
	const loaded = await new PptxHandler().load(bytes);
	return {
		// [0] is the slide's own group transform.
		xfrms: (xml.match(/<a:xfrm\b[^>]*>/gu) ?? []).slice(1),
		elements: loaded.slides[0]?.elements ?? [],
	};
}

const flips = (element: PptxElement | undefined) => [
	element?.flipHorizontal ?? false,
	element?.flipVertical ?? false,
];

describe('element flips', () => {
	it('writes and reads flips on shapes, text boxes, pictures and connectors', async () => {
		const { xfrms, elements } = await roundTrip((s) =>
			s
				.addShape('rtTriangle', { flipVertical: true })
				.addText('t', { flipHorizontal: true })
				.addImage(PNG, { flipVertical: true, flipHorizontal: true })
				.addConnector({ flipVertical: true })
				.addShape('rect'),
		);
		expect(xfrms[0]).toContain('flipV="1"');
		expect(xfrms[0]).not.toContain('rot=');
		expect(xfrms[0]).not.toContain('flipH');
		expect(xfrms[1]).toContain('flipH="1"');
		expect(xfrms[2]).toMatch(/flipH="1"[^>]*flipV="1"|flipV="1"[^>]*flipH="1"/u);
		expect(xfrms[3]).toContain('flipV="1"');
		expect(xfrms[4]).not.toContain('flip');
		expect(elements.map(flips)).toEqual([
			[false, true],
			[true, false],
			[true, true],
			[false, true],
			[false, false],
		]);
	});

	it('writes nothing for false flips', async () => {
		const { xfrms, elements } = await roundTrip((s) =>
			s.addShape('rect', { flipHorizontal: false, flipVertical: false }),
		);
		expect(xfrms[0]).not.toContain('flip');
		expect(flips(elements[0])).toEqual([false, false]);
	});

	it('keeps a rotation beside a flip', async () => {
		const { xfrms, elements } = await roundTrip((s) =>
			s.addShape('rtTriangle', { rotation: 30, flipHorizontal: true }),
		);
		expect(xfrms[0]).toContain('rot="1800000"');
		expect(xfrms[0]).toContain('flipH="1"');
		expect(elements[0]?.rotation).toBe(30);
		expect(flips(elements[0])).toEqual([true, false]);
	});

	it('writes and reads a flipped group', async () => {
		const { elements } = await roundTrip((s) =>
			s.addGroup([createShapeElement('rtTriangle')], { flipHorizontal: true }),
		);
		expect(elements[0]?.type).toBe('group');
		expect(flips(elements[0])).toEqual([true, false]);
	});

	it('sets the flips of a media element', () => {
		const el = createMediaElement('video', 'media/clip.mp4', { flipVertical: true });
		expect(flips(el)).toEqual([false, true]);
	});
});

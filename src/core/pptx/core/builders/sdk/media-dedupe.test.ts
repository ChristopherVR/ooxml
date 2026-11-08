/**
 * New pictures with the same bytes share one media part (and, on one slide,
 * one relationship); each keeps its own picture properties, and pictures
 * loaded from the deck keep their own parts.
 */
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';

import { PptxHandler } from '../../PptxHandler';
import type { ImagePptxElement, PptxElement } from '../../types';

const RED =
	'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';
const OTHER =
	'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==';

const mediaParts = (zip: JSZip) => Object.keys(zip.files).filter((n) => n.startsWith('ppt/media/'));
const embeds = (xml: string) => [...xml.matchAll(/r:embed="([^"]+)"/gu)].map((m) => m[1]);
const pictures = (elements: PptxElement[]) =>
	elements.filter((el) => el.type === 'image' || el.type === 'picture') as ImagePptxElement[];

async function read(zip: JSZip, name: string): Promise<string> {
	return (await zip.file(name)?.async('string')) ?? '';
}

describe('media de-duplication', () => {
	it('writes one part per distinct picture, across slides, and reads every picture back', async () => {
		const { handler, createSlide } = await PptxHandler.create();
		const slides = [
			createSlide('Blank').addImage(RED).addImage(RED).addImage(OTHER).build(),
			createSlide('Blank').addImage(RED).build(),
		];
		const bytes = await handler.save(slides);
		const zip = await JSZip.loadAsync(bytes);
		expect(mediaParts(zip)).toHaveLength(2);

		const ids = embeds(await read(zip, 'ppt/slides/slide1.xml'));
		expect(ids).toHaveLength(3);
		expect(ids[0]).toBe(ids[1]);
		expect(ids[2]).not.toBe(ids[0]);
		const rels1 = await read(zip, 'ppt/slides/_rels/slide1.xml.rels');
		expect(rels1.match(/relationships\/image"/gu)).toHaveLength(2);
		const rels2 = await read(zip, 'ppt/slides/_rels/slide2.xml.rels');
		const target1 = new RegExp(`Id="${ids[0]}"[^>]*Target="([^"]+)"`, 'u').exec(rels1)?.[1];
		expect(rels2).toContain(`Target="${target1}"`);

		const data = await new PptxHandler().load(bytes);
		const first = pictures(data.slides[0]?.elements ?? []);
		expect(first).toHaveLength(3);
		expect(first[0]?.imagePath).toBe(first[1]?.imagePath);
		expect(first[2]?.imagePath).not.toBe(first[0]?.imagePath);
		expect(pictures(data.slides[1]?.elements ?? [])[0]?.imagePath).toBe(first[0]?.imagePath);
	});

	it('keeps each picture its own crop and size when they share a part', async () => {
		const { handler, createSlide } = await PptxHandler.create();
		const slide = createSlide('Blank')
			.addImage(RED, { width: 50, height: 50, cropLeft: 0.25 })
			.addImage(RED, { x: 300, width: 120, height: 80 })
			.build();
		const bytes = await handler.save([slide]);
		const zip = await JSZip.loadAsync(bytes);
		expect(mediaParts(zip)).toHaveLength(1);
		const [a, b] = pictures((await new PptxHandler().load(bytes)).slides[0]?.elements ?? []);
		expect(a?.cropLeft).toBeCloseTo(0.25, 5);
		expect(b?.cropLeft ?? 0).toBe(0);
		expect([a?.width, b?.width]).toEqual([50, 120]);
	});

	it('gives a new picture its own part next to an identical loaded one', async () => {
		const { handler, createSlide } = await PptxHandler.create();
		const first = await handler.save([createSlide('Blank').addImage(RED).build()]);
		const reloaded = new PptxHandler();
		const data = await reloaded.load(first);
		const loadedPath = pictures(data.slides[0]?.elements ?? [])[0]?.imagePath;
		const slides = [...data.slides];
		const { createSlide: createMore } = await PptxHandler.create();
		slides.push({ ...createMore('Blank').addImage(RED).build(), id: 'new-slide' });
		const zip = await JSZip.loadAsync(await reloaded.save(slides));
		expect(mediaParts(zip)).toHaveLength(2);
		expect(mediaParts(zip)).toContain(loadedPath);
	});

	it('writes a part per picture again when every picture is different', async () => {
		const { handler, createSlide } = await PptxHandler.create();
		const zip = await JSZip.loadAsync(
			await handler.save([createSlide('Blank').addImage(RED).addImage(OTHER).build()]),
		);
		expect(mediaParts(zip)).toHaveLength(2);
		const ids = embeds(await read(zip, 'ppt/slides/slide1.xml'));
		expect(new Set(ids).size).toBe(2);
	});
});

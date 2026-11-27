import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';

import { PptxHandler } from '../../PptxHandler';
import type { PptxImageLikeElement, SmartArtPptxElement } from '../../types';
import { listPictureLayouts, pictureLayoutSmartArt } from './picture-layout';

const PNG =
	'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';

const picture = (id: string, x: number, y: number): PptxImageLikeElement =>
	({
		type: 'picture',
		id,
		x,
		y,
		width: 120,
		height: 90,
		imageData: PNG,
	}) as PptxImageLikeElement;

const layoutId = (title: string) => {
	const entry = listPictureLayouts().find((item) => item.title === title);
	if (!entry) throw new Error(`no picture layout ${title}`);
	return entry.id;
};

describe('Picture Layout', () => {
	it('offers the built-in picture layouts only', () => {
		const titles = listPictureLayouts().map((entry) => entry.title);
		expect(titles).toContain('Bending Picture Caption');
		expect(titles).toContain('Picture Accent List');
		expect(titles).not.toContain('Basic Block List');
	});

	it('puts each picture, in reading order, into a node of the chosen layout', async () => {
		const result = await pictureLayoutSmartArt(
			[picture('b', 300, 40), picture('a', 40, 40), picture('c', 40, 200)],
			layoutId('Picture Grid'),
			{ id: 'sa-1' },
		);
		if (typeof result === 'string') throw new Error(result);
		expect(result).toMatchObject({ type: 'smartArt', id: 'sa-1', x: 40, y: 40, width: 380 });
		const data = result.smartArtData!;
		expect(data.builtinLayoutId).toBe(layoutId('Picture Grid'));
		expect(data.nodes.map((node) => node.text)).toEqual(['', '', '']);
		const pictures = data.pictureLayoutSource?.pictures ?? [];
		expect(pictures.map((item) => item.nodeId)).toEqual(data.nodes.map((node) => node.id));
		// The drawing places the pictures in the layout's placeholders now, before any save.
		expect(data.drawingShapes?.filter((shape) => shape.fillImageUrl === PNG)).toHaveLength(3);
	});

	it('refuses pictures it cannot read and layouts it does not know', async () => {
		const unreadable = { ...picture('a', 0, 0), imageData: undefined } as PptxImageLikeElement;
		expect(await pictureLayoutSmartArt([unreadable], layoutId('Picture Grid'), { id: 's' })).toBe(
			'pictureUnavailable',
		);
		expect(await pictureLayoutSmartArt([picture('a', 0, 0)], 'urn:nope', { id: 's' })).toBe(
			'unknownLayout',
		);
		expect(await pictureLayoutSmartArt([], layoutId('Picture Grid'), { id: 's' })).toBe(
			'noPictures',
		);
	});

	it('saves the built-in layout, the pictures as media and their relationships', async () => {
		const { handler, createSlide } = await PptxHandler.create({ title: 'Pictures' });
		const slide = createSlide('Blank').build();
		const smartArt = (await pictureLayoutSmartArt(
			[picture('a', 40, 40), picture('b', 300, 40)],
			layoutId('Bending Picture Caption'),
			{ id: 'sa-1' },
		)) as SmartArtPptxElement;
		slide.elements.push(smartArt);
		const zip = await JSZip.loadAsync(await handler.save([slide]));
		const names = Object.keys(zip.files);
		const layout = names.find((name) => /^ppt\/diagrams\/layout\d+\.xml$/u.test(name));
		expect(await zip.file(layout!)?.async('string')).toContain(
			'urn:microsoft.com/office/officeart/2008/layout/BendingPictureCaption',
		);
		const dataRels = names.find((name) => /^ppt\/diagrams\/_rels\/data\d+\.xml\.rels$/u.test(name));
		const rels = (await zip.file(dataRels!)?.async('string')) ?? '';
		const media = [...rels.matchAll(/Target="\/(ppt\/media\/[^"]+)"/gu)].map((match) => match[1]);
		expect(media.length).toBeGreaterThan(0);
		for (const part of media) expect(zip.file(part!)).not.toBeNull();
		expect(names.some((name) => name.includes('ooxml-picture-layout-'))).toBe(false);

		// Reloaded, the diagram still names a picture for each node.
		const reloaded = await new PptxHandler().load(
			(await zip.generateAsync({ type: 'arraybuffer' })) as ArrayBuffer,
		);
		const loaded = reloaded.slides[0]?.elements.find(
			(element): element is SmartArtPptxElement => element.type === 'smartArt',
		);
		expect(loaded?.smartArtData?.pictureLayoutSource?.pictures).toHaveLength(2);
	});
});

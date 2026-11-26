import JSZip from 'jszip';
import { expect, it } from 'vitest';

import { PresentationBuilder } from '../../core/builders/sdk/PresentationBuilder';
import { PptxHandler } from '../../core/PptxHandler';
import { collectSlideInk, convertSlideInk, lassoSlideInk, replaceSlideInk } from '../../editor/ink';

it('round-trips captured timing and converts only the lassoed stroke to a preset shape', async () => {
	const { handler, data, createSlide } = await PresentationBuilder.create();
	const slide = createSlide('Blank').build();
	slide.elements.push({
		id: 'captured-ink',
		type: 'ink',
		x: 100,
		y: 80,
		width: 200,
		height: 120,
		inkPaths: ['M 0 0 L 70 0 L 70 40 L 0 40 L 0 0', 'M 120 60 L 170 90'],
		inkColors: ['#336699', '#CC3300'],
		inkPointTimestamps: [
			[100, 160, 220, 280, 340],
			[580, 720],
		],
	});
	data.slides.push(slide);
	const first = await handler.save(data.slides);
	const fresh = new PptxHandler();
	const reloaded = await fresh.load(first);
	const active = reloaded.slides[0];
	const strokes = collectSlideInk(active.elements);
	expect(strokes.map((stroke) => stroke.style.pointTimestamps)).toStrictEqual([
		[100, 160, 220, 280, 340],
		[580, 720],
	]);
	const selected = lassoSlideInk(strokes, [
		{ x: 90, y: 70 },
		{ x: 180, y: 70 },
		{ x: 180, y: 130 },
		{ x: 90, y: 130 },
	]);
	expect(selected).toHaveLength(1);
	const shape = convertSlideInk(selected, 'converted');
	expect(shape).toMatchObject({ type: 'shape', shapeType: 'rect', x: 100, y: 80 });
	if (!shape) throw new Error('expected a preset shape');
	active.elements = replaceSlideInk(active.elements, selected, shape);
	const second = await fresh.save(reloaded.slides);
	const zip = await JSZip.loadAsync(second);
	const slideXml = await zip.file('ppt/slides/slide1.xml')?.async('string');
	expect(slideXml).toContain('<a:prstGeom prst="rect"');
	expect(slideXml).toContain('<p:contentPart');
	const final = await new PptxHandler().load(second);
	const remaining = collectSlideInk(final.slides[0].elements);
	expect(remaining).toHaveLength(1);
	expect(remaining[0].style).toMatchObject({ color: '#CC3300', pointTimestamps: [580, 720] });
});

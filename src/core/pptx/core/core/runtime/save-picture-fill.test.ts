/**
 * Picture fills of new shapes (`FillInput { type: 'image' }`) and the tile /
 * stretch layout of new pictures and picture fills (`a:tile`,
 * `a:stretch/a:fillRect`), saved through the public SDK and reloaded.
 */
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';

import { createGroupElement, createShapeElement } from '../../builders/sdk/ElementFactory';
import { PresentationBuilder } from '../../builders/sdk/PresentationBuilder';
import type { SlideBuilder } from '../../builders/sdk/SlideBuilder';
import { PptxHandler } from '../../PptxHandler';
import type { PptxElement, PptxImageLayout, ShapeStyle } from '../../types';

/** 1x1 red PNG. */
const PNG =
	'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';
const box = { x: 10, y: 10, width: 100, height: 80 };

interface Saved {
	xml: string;
	rels: string;
	zip: JSZip;
	warnings: string[];
	elements: Array<PptxElement & PptxImageLayout & { imagePath?: string; shapeStyle?: ShapeStyle }>;
}

/** Build one slide, save it, and return its XML, relationships, warnings and reloaded elements. */
async function roundTrip(build: (slide: SlideBuilder) => void): Promise<Saved> {
	const { handler, createSlide } = await PresentationBuilder.create();
	const slide = createSlide('Blank');
	build(slide);
	const bytes = await handler.save([slide.build()]);
	const zip = await JSZip.loadAsync(bytes);
	const raw = (await zip.file('ppt/slides/slide1.xml')?.async('string')) ?? '';
	// The writer spells empty elements `<x></x>`; compare them as `<x/>`.
	const xml = raw.replace(/<([\w:]+)([^<>]*)><\/\1>/gu, '<$1$2/>');
	const rels = (await zip.file('ppt/slides/_rels/slide1.xml.rels')?.async('string')) ?? '';
	const warnings = handler.getCompatibilityWarnings().map((w) => w.code);
	const loaded = await new PptxHandler().load(bytes);
	return {
		xml,
		rels,
		zip,
		warnings,
		elements: (loaded.slides[0]?.elements ?? []) as Saved['elements'],
	};
}

/** The `a:blipFill` inside the first `p:spPr`. */
function shapeBlipFill(xml: string): string {
	return /<p:spPr>.*?(<a:blipFill\b.*?<\/a:blipFill>)/su.exec(xml)?.[1] ?? '';
}

/** The relationship target of `rId`. */
function target(rels: string, rId: string): string | undefined {
	return new RegExp(`Id="${rId}"[^>]*Target="([^"]+)"`, 'u').exec(rels)?.[1];
}

describe('shape picture fill', () => {
	it('embeds the picture and writes a:blipFill with a full-frame stretch by default', async () => {
		const { xml, rels, zip, warnings, elements } = await roundTrip((s) =>
			s.addShape('ellipse', { ...box, fill: { type: 'image', url: PNG } }),
		);
		const fill = shapeBlipFill(xml);
		expect(fill).toMatch(
			/^<a:blipFill rotWithShape="1"><a:blip r:embed="(rId\d+)"\/><a:stretch><a:fillRect\/><\/a:stretch><\/a:blipFill>$/u,
		);
		const rId = /r:embed="(rId\d+)"/u.exec(fill)?.[1] ?? '';
		expect(rels).toMatch(new RegExp(`Id="${rId}" Type="[^"]+/relationships/image"`, 'u'));
		const media = target(rels, rId)?.replace('../', 'ppt/') ?? '';
		expect(await zip.file(media)?.async('base64')).toBe(PNG.split(',')[1]);
		expect(warnings).toEqual([]);
		// A shape with a picture fill loads as a picture element.
		expect(elements).toHaveLength(1);
		expect(elements[0]).toMatchObject({ type: 'picture', shapeType: 'ellipse', imagePath: media });
		expect(elements[0]?.shapeStyle?.fillMode).toBe('image');
	});

	it('fills a text box with a picture too', async () => {
		const { xml } = await roundTrip((s) =>
			s.addText('Caption', { ...box, fill: { type: 'image', url: PNG } }),
		);
		expect(shapeBlipFill(xml)).toContain('<a:stretch><a:fillRect/></a:stretch>');
	});

	it('writes the fill after the geometry and before the outline, as its only fill', async () => {
		const { xml } = await roundTrip((s) =>
			s.addShape('rect', {
				...box,
				fill: { type: 'image', url: PNG },
				stroke: { color: '#000000', width: 2 },
			}),
		);
		const spPr = /<p:spPr>(.*?)<\/p:spPr>/su.exec(xml)?.[1] ?? '';
		const children = [...spPr.matchAll(/<(a:\w+)/gu)].map((m) => m[1]);
		const top = children.filter((name) =>
			['a:xfrm', 'a:prstGeom', 'a:blipFill', 'a:ln', 'a:solidFill', 'a:noFill'].includes(
				name ?? '',
			),
		);
		// `a:solidFill` here is the outline's own, inside `a:ln`.
		expect(top.slice(0, 4)).toEqual(['a:xfrm', 'a:prstGeom', 'a:blipFill', 'a:ln']);
		expect(spPr.replace(/<a:ln\b.*<\/a:ln>/su, '')).not.toMatch(/<a:(solidFill|noFill|gradFill)/u);
	});

	it('writes a tile with every attribute given', async () => {
		const { xml, elements } = await roundTrip((s) =>
			s.addShape('rect', {
				...box,
				fill: {
					type: 'image',
					url: PNG,
					tile: { offsetX: 10, offsetY: -4, scaleX: 0.5, scaleY: 2, flip: 'xy', alignment: 'ctr' },
				},
			}),
		);
		const tile = /<a:tile\b[^>]*\/>/u.exec(shapeBlipFill(xml))?.[0] ?? '';
		expect(tile).toContain('tx="95250"');
		expect(tile).toContain('ty="-38100"');
		expect(tile).toContain('sx="50000"');
		expect(tile).toContain('sy="200000"');
		expect(tile).toContain('flip="xy"');
		expect(tile).toContain('algn="ctr"');
		expect(shapeBlipFill(xml)).not.toContain('a:stretch');
		expect(elements[0]).toMatchObject({
			tileOffsetX: 10,
			tileOffsetY: -4,
			tileScaleX: 0.5,
			tileScaleY: 2,
			tileFlip: 'xy',
			tileAlignment: 'ctr',
		});
	});

	it("writes PowerPoint's tile defaults for mode 'tile' alone", async () => {
		const { xml, elements } = await roundTrip((s) =>
			s.addShape('rect', { ...box, fill: { type: 'image', url: PNG, mode: 'tile' } }),
		);
		const tile = /<a:tile\b[^>]*\/>/u.exec(shapeBlipFill(xml))?.[0] ?? '';
		for (const attribute of [
			'tx="0"',
			'ty="0"',
			'sx="100000"',
			'sy="100000"',
			'flip="none"',
			'algn="tl"',
		]) {
			expect(tile).toContain(attribute);
		}
		expect(elements[0]).toMatchObject({
			tileScaleX: 1,
			tileScaleY: 1,
			tileFlip: 'none',
			tileAlignment: 'tl',
		});
	});

	it('writes stretch insets, negative ones included', async () => {
		const { xml, elements } = await roundTrip((s) =>
			s.addShape('rect', {
				...box,
				fill: { type: 'image', url: PNG, fillRect: { l: 0.1, t: -0.25, r: 0, b: 0.3 } },
			}),
		);
		expect(shapeBlipFill(xml)).toContain(
			'<a:stretch><a:fillRect l="10000" t="-25000" r="0" b="30000"/></a:stretch>',
		);
		expect(elements[0]).toMatchObject({
			fillRectLeft: 0.1,
			fillRectTop: -0.25,
			fillRectRight: 0,
			fillRectBottom: 0.3,
		});
	});

	it('keeps the extremes the loader reads back', async () => {
		const { elements } = await roundTrip((s) =>
			s.addShape('rect', {
				...box,
				fill: { type: 'image', url: PNG, fillRect: { l: -10, t: -10, r: 10, b: 10 } },
			}),
		);
		expect(elements[0]).toMatchObject({
			fillRectLeft: -10,
			fillRectTop: -10,
			fillRectRight: 10,
			fillRectBottom: 10,
		});
	});

	it('rounds a sub-pixel tile offset to whole EMU', async () => {
		const { xml, elements } = await roundTrip((s) =>
			s.addShape('rect', { ...box, fill: { type: 'image', url: PNG, tile: { offsetX: 0.5 } } }),
		);
		expect(shapeBlipFill(xml)).toContain('tx="4763"');
		expect(elements[0]?.tileOffsetX).toBeCloseTo(0.5, 3);
	});

	it('gives each filled shape its own relationship and media part', async () => {
		const { xml, rels } = await roundTrip((s) => {
			s.addShape('rect', { ...box, fill: { type: 'image', url: PNG } });
			s.addImage(PNG, box);
			s.addShape('rect', { ...box, fill: { type: 'image', url: PNG, mode: 'tile' } });
		});
		const ids = [...xml.matchAll(/r:embed="(rId\d+)"/gu)].map((m) => m[1] ?? '');
		expect(new Set(ids).size).toBe(3);
		const targets = ids.map((id) => target(rels, id));
		expect(new Set(targets).size).toBe(3);
		expect(targets.every((t) => t?.startsWith('../media/'))).toBe(true);
	});

	it('fills a shape inside a group', async () => {
		const child = createShapeElement('rect', { ...box, fill: { type: 'image', url: PNG } });
		const { xml, warnings } = await roundTrip((s) => s.addElement(createGroupElement([child])));
		expect(xml).toMatch(/<p:grpSp>.*<a:blipFill rotWithShape="1"><a:blip r:embed="rId\d+"\/>/su);
		expect(warnings).toEqual([]);
	});

	it('writes rotWithShape="0" when the model turns it off', async () => {
		const shape = createShapeElement('rect', { ...box, fill: { type: 'image', url: PNG } });
		shape.shapeStyle = { ...shape.shapeStyle, fillImageRotWithShape: false };
		const { xml } = await roundTrip((s) => s.addElement(shape));
		expect(shapeBlipFill(xml)).toMatch(/^<a:blipFill rotWithShape="0">/u);
	});

	it('warns instead of dropping a picture fill it cannot embed', async () => {
		const shape = createShapeElement('rect', { ...box, fill: { type: 'image', url: PNG } });
		shape.shapeStyle = { ...shape.shapeStyle, fillImageUrl: 'https://example.com/a.png' };
		const { xml, warnings } = await roundTrip((s) => s.addElement(shape));
		expect(xml).not.toContain('a:blipFill');
		expect(warnings).toContain('SAVE_SHAPE_PICTURE_FILL_UNSUPPORTED');
	});

	it('warns and stretches when the model holds a tile the file cannot', async () => {
		const shape = createShapeElement('rect', { ...box, fill: { type: 'image', url: PNG } });
		shape.shapeStyle = { ...shape.shapeStyle, fillImageLayout: { tileScaleX: 0 } };
		const { xml, warnings } = await roundTrip((s) => s.addElement(shape));
		expect(shapeBlipFill(xml)).toContain('<a:stretch><a:fillRect/></a:stretch>');
		expect(shapeBlipFill(xml)).not.toContain('a:tile');
		expect(warnings).toContain('SAVE_PICTURE_LAYOUT_INVALID');
	});
});

describe('picture tile and stretch', () => {
	/** The `p:blipFill` of the first picture. */
	const picBlipFill = (xml: string) => /<p:blipFill>.*?<\/p:blipFill>/su.exec(xml)?.[0] ?? '';

	it('keeps the full-frame stretch without options', async () => {
		const { xml } = await roundTrip((s) => s.addImage(PNG, box));
		expect(picBlipFill(xml)).toMatch(
			/^<p:blipFill><a:blip r:embed="rId\d+"\/><a:stretch><a:fillRect\/><\/a:stretch><\/p:blipFill>$/u,
		);
	});

	it('tiles a picture', async () => {
		const { xml, elements } = await roundTrip((s) =>
			s.addImage(PNG, { ...box, tile: { scaleX: 0.25, scaleY: 0.25, alignment: 'br', flip: 'x' } }),
		);
		const fill = picBlipFill(xml);
		expect(fill).toMatch(/<a:blip r:embed="rId\d+"\/><a:tile\b/u);
		expect(fill).not.toContain('a:stretch');
		for (const attribute of [
			'sx="25000"',
			'sy="25000"',
			'algn="br"',
			'flip="x"',
			'tx="0"',
			'ty="0"',
		]) {
			expect(fill).toContain(attribute);
		}
		expect(elements[0]).toMatchObject({
			type: 'picture',
			tileScaleX: 0.25,
			tileScaleY: 0.25,
			tileAlignment: 'br',
			tileFlip: 'x',
			tileOffsetX: 0,
			tileOffsetY: 0,
		});
	});

	it('stretches a picture into an inset rectangle', async () => {
		const { xml, elements } = await roundTrip((s) =>
			s.addImage(PNG, { ...box, fillRect: { l: 0.25, t: 0, r: -0.5, b: 0.125 } }),
		);
		expect(picBlipFill(xml)).toContain(
			'<a:stretch><a:fillRect l="25000" t="0" r="-50000" b="12500"/></a:stretch>',
		);
		expect(elements[0]).toMatchObject({
			fillRectLeft: 0.25,
			fillRectTop: 0,
			fillRectRight: -0.5,
			fillRectBottom: 0.125,
		});
	});

	it('keeps a:srcRect before a:stretch with a crop', async () => {
		const { xml } = await roundTrip((s) =>
			s.addImage(PNG, { ...box, cropLeft: 0.1, fillRect: { l: 0.2, t: 0, r: 0, b: 0 } }),
		);
		expect(picBlipFill(xml)).toMatch(/<a:blip\b[^>]*\/><a:srcRect\b[^>]*\/><a:stretch>/u);
	});

	it('keeps a:tile after a blip that carries effects', async () => {
		const { xml } = await roundTrip((s) => {
			s.addImage(PNG, { ...box, tile: {} });
			const picture = s.getLastElement() as PptxElement & { imageEffects?: object };
			picture.imageEffects = { grayscale: true };
		});
		expect(picBlipFill(xml)).toMatch(/<a:blip\b[^>]*><a:grayscl\/><\/a:blip><a:tile\b/u);
	});
});

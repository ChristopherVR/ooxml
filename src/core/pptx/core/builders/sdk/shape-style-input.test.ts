/**
 * Shape fill, outline and effect inputs on NEW elements, driven through the
 * public SDK and checked in the saved slide XML: path gradients and pattern
 * foregrounds (`FillInput`), exact dashes, compound lines, alignment, miter
 * limits and line ends (`StrokeInput`), and inner shadow, glow, soft edge and
 * reflection (`ShapeOptions` / `TextOptions`).
 */
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';

import type { PptxElement } from '../../types/elements';
import { createConnectorElement, createShapeElement } from './ElementFactory';
import { PresentationBuilder } from './PresentationBuilder';
import { ShapeBuilder } from './ShapeBuilder';
import type { ShapeOptions } from './types';

/** Save one slide holding `elements` and return its slide XML. */
async function savedSlideXml(elements: PptxElement[]): Promise<string> {
	const { handler, data, createSlide } = await PresentationBuilder.create();
	const slide = createSlide('Blank').build();
	slide.elements.push(...elements);
	const bytes = await handler.save([slide]);
	const zip = await JSZip.loadAsync(bytes);
	const xml = await zip.file('ppt/slides/slide1.xml')?.async('string');
	expect(xml, `slides: ${data.slides.length}`).toBeDefined();
	// The writer spells empty elements `<x></x>`; compare them as `<x/>`.
	return (xml ?? '').replace(/<([\w:]+)([^<>]*)><\/\1>/gu, '<$1$2/>');
}

async function shapeXml(options: ShapeOptions): Promise<string> {
	return savedSlideXml([
		createShapeElement('rect', { x: 10, y: 10, width: 100, height: 60, ...options }),
	]);
}

/** The `a:ln` element of the first shape, without its attributes' order mattering. */
function lineXml(xml: string): string {
	const match = /<a:ln\b[^>]*>[\s\S]*?<\/a:ln>/u.exec(xml);
	expect(match).not.toBeNull();
	return match?.[0] ?? '';
}

describe('stroke input', () => {
	it('writes an exact dash pattern, compound type and alignment', async () => {
		const line = lineXml(
			await shapeXml({
				stroke: { color: '#000000', width: 4, dash: [6, 2], compound: 'dbl', align: 'in' },
			}),
		);
		expect(line).toMatch(/<a:ln\b[^>]*\bw="38100"/u);
		expect(line).toMatch(/<a:ln\b[^>]*\bcmpd="dbl"/u);
		expect(line).toMatch(/<a:ln\b[^>]*\balgn="in"/u);
		expect(line).toContain('<a:custDash><a:ds d="150000" sp="50000"/></a:custDash>');
		expect(line).not.toContain('a:prstDash');
	});

	it('repeats an odd dash list once, as SVG does', async () => {
		const line = lineXml(
			await shapeXml({ stroke: { color: '#000000', width: 2, dash: [4, 2, 1] } }),
		);
		expect(line).toContain(
			'<a:custDash><a:ds d="200000" sp="100000"/><a:ds d="50000" sp="200000"/><a:ds d="100000" sp="50000"/></a:custDash>',
		);
	});

	it('writes a miter limit as a ratio of the line width', async () => {
		const line = lineXml(await shapeXml({ stroke: { color: '#000000', width: 1, miterLimit: 4 } }));
		expect(line).toContain('<a:miter lim="400000"/>');
	});

	it('writes line ends on a shape outline in schema order', async () => {
		const line = lineXml(
			await shapeXml({
				stroke: {
					color: '#FF0000',
					width: 2,
					join: 'round',
					headEnd: { type: 'triangle', width: 'lg' },
					tailEnd: { type: 'oval', length: 'sm' },
				},
			}),
		);
		const children = [...line.matchAll(/<(a:[A-Za-z]+)/gu)].map((m) => m[1]).slice(1);
		expect(children).toEqual(['a:solidFill', 'a:srgbClr', 'a:round', 'a:headEnd', 'a:tailEnd']);
		expect(line).toContain('<a:headEnd type="triangle" w="lg"/>');
		expect(line).toContain('<a:tailEnd type="oval" len="sm"/>');
	});

	it('takes connector ends from the stroke when no arrow option is given', async () => {
		const connector = createConnectorElement({
			x: 0,
			y: 0,
			width: 100,
			height: 0,
			stroke: { color: '#000000', width: 1, tailEnd: { type: 'stealth' } },
		});
		expect(connector.shapeStyle?.connectorEndArrow).toBe('stealth');
		expect(connector.shapeStyle?.connectorStartArrow).toBe('none');
		const withArrow = createConnectorElement({
			endArrow: 'arrow',
			stroke: { tailEnd: { type: 'stealth' } },
		});
		expect(withArrow.shapeStyle?.connectorEndArrow).toBe('arrow');
	});
});

describe('fill input', () => {
	it('writes a path gradient with its focus rectangle', async () => {
		const xml = await shapeXml({
			fill: {
				type: 'gradient',
				path: 'circle',
				focus: { l: 0.5, t: 0.5, r: 0.5, b: 0.5 },
				stops: [
					{ color: '#FFFFFF', position: 0 },
					{ color: '#1F4E79', position: 100 },
				],
			},
		});
		expect(xml).toContain(
			'<a:path path="circle"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path>',
		);
		expect(xml).not.toContain('<a:lin ');
	});

	it('writes a rect path, tile rectangle and rotWithShape', async () => {
		const xml = await shapeXml({
			fill: {
				type: 'gradient',
				path: 'rect',
				tile: { l: 0, t: 0, r: -1, b: -1 },
				rotWithShape: false,
				stops: [
					{ color: '#FFFFFF', position: 0 },
					{ color: '#000000', position: 100 },
				],
			},
		});
		expect(xml).toMatch(/<a:gradFill\b[^>]*\brotWithShape="0"/u);
		expect(xml).toContain('<a:path path="rect"');
		expect(xml).toContain('<a:tileRect l="0" t="0" r="-100000" b="-100000"/>');
	});

	it('writes a linear gradient that does not scale with the shape', async () => {
		const xml = await shapeXml({
			fill: {
				type: 'gradient',
				angle: 45,
				scaled: false,
				stops: [
					{ color: '#FFFFFF', position: 0 },
					{ color: '#000000', position: 100 },
				],
			},
		});
		expect(xml).toContain('<a:lin ang="2700000" scaled="0"/>');
	});

	it('writes the pattern foreground and background colours', async () => {
		const xml = await shapeXml({
			fill: { type: 'pattern', preset: 'dkDnDiag', foreground: '#C00000', background: '#FFFFFF' },
		});
		expect(xml).toContain(
			'<a:pattFill prst="dkDnDiag"><a:fgClr><a:srgbClr val="C00000"/></a:fgClr><a:bgClr><a:srgbClr val="FFFFFF"/></a:bgClr></a:pattFill>',
		);
	});
});

describe('effect input', () => {
	it('writes glow, inner shadow and soft edge in effect-list order', async () => {
		const xml = await shapeXml({
			glow: { color: '#FFC000', radius: 8, opacity: 0.6 },
			innerShadow: { offsetX: -4, offsetY: -4, blur: 5, opacity: 0.5 },
			softEdge: 4,
		});
		const effects = /<a:effectLst>([\s\S]*?)<\/a:effectLst>/u.exec(xml)?.[1] ?? '';
		const children = [...effects.matchAll(/<(a:(?:glow|innerShdw|softEdge|outerShdw))\b/gu)].map(
			(m) => m[1],
		);
		expect(children).toEqual(['a:glow', 'a:innerShdw', 'a:softEdge']);
		expect(effects).toMatch(/<a:glow rad="76200"><a:srgbClr val="FFC000"><a:alpha val="60000"\/>/u);
		expect(effects).toMatch(/<a:innerShdw\b[^>]*\bdir="13500000"/u);
		expect(effects).toContain('<a:softEdge rad="38100"/>');
	});

	it('writes a mirrored reflection', async () => {
		const xml = await shapeXml({ reflection: { blur: 1, distance: 2, startOpacity: 0.4 } });
		const reflection = /<a:reflection\b[^>]*\/>/u.exec(xml)?.[0] ?? '';
		expect(reflection).toContain('blurRad="9525"');
		expect(reflection).toContain('dist="19050"');
		expect(reflection).toContain('stA="40000"');
		expect(reflection).toContain('endA="0"');
		expect(reflection).toContain('endPos="90000"');
		expect(reflection).toContain('dir="5400000"');
		expect(reflection).toContain('sy="-100000"');
		expect(reflection).toContain('algn="bl"');
	});

	it('accepts the effects through the fluent shape builder', () => {
		const shape = ShapeBuilder.create('ellipse')
			.glow({ color: '#00B0F0', radius: 4 })
			.softEdge(2)
			.innerShadow({ color: '#333333' })
			.reflection({})
			.build();
		expect(shape.shapeStyle).toMatchObject({
			glowColor: '#00B0F0',
			glowRadius: 4,
			softEdgeRadius: 2,
			innerShadowColor: '#333333',
			reflectionStartOpacity: 0.5,
		});
	});
});

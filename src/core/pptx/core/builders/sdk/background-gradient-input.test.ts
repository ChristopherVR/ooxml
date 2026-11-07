/**
 * `SlideBuilder.setBackground({ type: 'gradient' })`: the structured gradient
 * is saved as `p:bg/p:bgPr/a:gradFill` by the shape gradient writer and
 * reloads as a gradient background; invalid input is refused.
 */
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';

import { PptxHandler } from '../../PptxHandler';
import type { PptxSlide } from '../../types';
import { PresentationBuilder } from './PresentationBuilder';
import type { SlideBuilder } from './SlideBuilder';
import type { BackgroundInput } from './types';

type GradientInput = Extract<BackgroundInput, { type: 'gradient' }>;
const RED_BLUE: GradientInput['stops'] = [
	{ color: '#FF0000', position: 0 },
	{ color: '#0000FF', position: 1 },
];
const PNG =
	'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';

interface Saved {
	xml: string;
	bg: string;
	warnings: string[];
	reloaded: PptxSlide | undefined;
	bytes: Uint8Array;
}

/** Build one slide, save it, and return its XML, `p:bg`, warnings and reloaded slide. */
async function roundTrip(build: (slide: SlideBuilder) => void, edit?: (slide: PptxSlide) => void) {
	const { handler, createSlide } = await PresentationBuilder.create();
	const builder = createSlide('Blank');
	build(builder);
	const slide = builder.build();
	edit?.(slide);
	const bytes = await handler.save([slide]);
	const zip = await JSZip.loadAsync(bytes);
	const raw = (await zip.file('ppt/slides/slide1.xml')?.async('string')) ?? '';
	// The writer spells empty elements `<x></x>`; compare them as `<x/>`.
	const xml = raw.replace(/<([\w:]+)([^<>]*)><\/\1>/gu, '<$1$2/>');
	const bg = /<p:bg>.*?<\/p:bg>/su.exec(xml)?.[0] ?? '';
	const warnings = handler.getCompatibilityWarnings().map((w) => w.code);
	const reloaded = (await new PptxHandler().load(bytes)).slides[0];
	return { xml, bg, warnings, reloaded, bytes } satisfies Saved;
}

const gradient = (extra: Partial<GradientInput> = {}): BackgroundInput => ({
	type: 'gradient',
	stops: RED_BLUE,
	...extra,
});

describe('slide gradient background', () => {
	it('writes a linear gradient, top to bottom by default', async () => {
		const { xml, bg, warnings, reloaded } = await roundTrip((s) => s.setBackground(gradient()));
		expect(bg).toBe(
			'<p:bg><p:bgPr><a:gradFill><a:gsLst>' +
				'<a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs>' +
				'<a:gs pos="100000"><a:srgbClr val="0000FF"/></a:gs>' +
				'</a:gsLst><a:lin ang="5400000" scaled="1"/></a:gradFill><a:effectLst/></p:bgPr></p:bg>',
		);
		// CT_CommonSlideData: `p:bg` comes before `p:spTree`.
		expect(xml).toMatch(/<p:cSld><p:bg>/u);
		expect(warnings).toEqual([]);
		expect(reloaded?.backgroundGradient).toMatch(/^linear-gradient\(180(\.0+)?deg, /u);
		expect(reloaded?.backgroundGradient).toContain('#FF0000');
		expect(reloaded?.backgroundGradient).toContain('#0000FF');
	});

	it.each([
		[90, '0'],
		[0, '16200000'],
		[135, '2700000'],
		[270, '10800000'],
	])('turns the CSS angle %d into a:lin ang="%s"', async (angle, ang) => {
		const { bg } = await roundTrip((s) => s.setBackground(gradient({ angle })));
		expect(bg).toContain(`<a:lin ang="${ang}" scaled="1"/>`);
	});

	it('keeps fractional positions, stop opacity and lowercase colours', async () => {
		const { bg, reloaded } = await roundTrip((s) =>
			s.setBackground(
				gradient({
					stops: [
						{ color: '#ff0000', position: 0 },
						{ color: '00FF00', position: 0.333, opacity: 0.5 },
						{ color: '#0000FF', position: 1, opacity: 0 },
					],
				}),
			),
		);
		expect(bg).toContain('<a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs>');
		expect(bg).toContain('<a:gs pos="33300"><a:srgbClr val="00FF00"><a:alpha val="50000"/>');
		expect(bg).toContain('<a:gs pos="100000"><a:srgbClr val="0000FF"><a:alpha val="0"/>');
		// The loader's CSS builder keeps float noise (`33.300000000000004%`).
		expect(reloaded?.backgroundGradient).toMatch(/rgba\(0, 255, 0, 0\.5\) 33\.3\d*%/u);
	});

	it('writes a path gradient from the centre by default', async () => {
		const { bg, reloaded } = await roundTrip((s) => s.setBackground(gradient({ path: 'circle' })));
		expect(bg).toContain(
			'<a:path path="circle"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path>',
		);
		expect(bg).not.toContain('a:lin');
		expect(reloaded?.backgroundGradient).toMatch(/^radial-gradient\(circle at 50% 50%/u);
	});

	it('writes a rect path gradient from a corner focus', async () => {
		const { bg } = await roundTrip((s) =>
			s.setBackground(gradient({ path: 'rect', focus: { l: 1, t: 1, r: 0, b: 0 } })),
		);
		expect(bg).toContain(
			'<a:path path="rect"><a:fillToRect l="100000" t="100000" r="0" b="0"/></a:path>',
		);
	});

	it('wins over a solid colour set before, as the renderers paint it', async () => {
		const { bg } = await roundTrip((s) =>
			s.setBackground({ type: 'solid', color: '#00FF00' }).setBackground(gradient()),
		);
		expect(bg).toContain('<a:gradFill>');
		expect(bg).not.toContain('a:solidFill');
	});

	it('is replaced by a solid colour set after it', async () => {
		let built: PptxSlide | undefined;
		const { bg } = await roundTrip(
			(s) => s.setBackground(gradient()).setBackground({ type: 'solid', color: '#00FF00' }),
			(slide) => {
				built = slide;
			},
		);
		expect(bg).toContain('<a:solidFill><a:srgbClr val="00FF00"/></a:solidFill>');
		expect(bg).not.toContain('a:gradFill');
		expect(built?.backgroundGradient).toBeUndefined();
		expect(built?.backgroundGradientFill).toBeUndefined();
	});

	it('gives way to a background picture', async () => {
		const { bg } = await roundTrip((s) =>
			s.setBackground(gradient()).setBackground({ type: 'image', source: PNG }),
		);
		expect(bg).toContain('<a:blipFill>');
		expect(bg).not.toContain('a:gradFill');
	});

	it('survives a second save of the reloaded deck', async () => {
		const { bytes } = await roundTrip((s) => s.setBackground(gradient({ angle: 90 })));
		const handler = new PptxHandler();
		const data = await handler.load(bytes);
		const again = await JSZip.loadAsync(await handler.save(data.slides));
		const xml = (await again.file('ppt/slides/slide1.xml')?.async('string')) ?? '';
		expect(xml).toMatch(/<p:bg><p:bgPr><a:gradFill>.*<a:lin ang="0"/su);
	});

	it('ignores a structured gradient whose CSS was replaced since, and warns', async () => {
		const { bg, warnings } = await roundTrip(
			(s) => s.setBackground(gradient()),
			(slide) => {
				slide.backgroundGradient = 'linear-gradient(45deg, #000000, #FFFFFF)';
			},
		);
		expect(bg).toBe('');
		expect(warnings).toContain('SAVE_BACKGROUND_GRADIENT_UNSUPPORTED');
	});

	it('warns instead of silently dropping a CSS-only gradient background', async () => {
		const { bg, warnings } = await roundTrip(
			() => undefined,
			(slide) => {
				slide.backgroundGradient = 'linear-gradient(180deg, #000000 0%, #FFFFFF 100%)';
			},
		);
		expect(bg).toBe('');
		expect(warnings).toContain('SAVE_BACKGROUND_GRADIENT_UNSUPPORTED');
	});

	it.each<[string, Partial<GradientInput>]>([
		['one stop', { stops: [{ color: '#FF0000', position: 0 }] }],
		['no stops', { stops: [] }],
		['a three-digit colour', { stops: [{ color: '#F00', position: 0 }, RED_BLUE[1]!] }],
		['a named colour', { stops: [{ color: 'red', position: 0 }, RED_BLUE[1]!] }],
		['a position above 1', { stops: [RED_BLUE[0]!, { color: '#0000FF', position: 100 }] }],
		['a negative position', { stops: [{ color: '#FF0000', position: -0.1 }, RED_BLUE[1]!] }],
		[
			'an opacity above 1',
			{ stops: [{ color: '#FF0000', position: 0, opacity: 2 }, RED_BLUE[1]!] },
		],
		['a NaN angle', { angle: Number.NaN }],
		['an unknown path', { path: 'diamond' as 'rect' }],
		['an angle with a path', { path: 'circle', angle: 90 }],
		['a focus without a path', { focus: { l: 0, t: 0, r: 0, b: 0 } }],
		['a focus with an infinite inset', { path: 'rect', focus: { l: Infinity, t: 0, r: 0, b: 0 } }],
	])('refuses %s', async (_name, extra) => {
		const { createSlide } = await PresentationBuilder.create();
		expect(() => createSlide('Blank').setBackground(gradient(extra))).toThrow(RangeError);
	});
});

import type { PptxElement, PptxSlide } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import { buildMorphTransitionPlan } from './morph-plan';
import type { MorphTextMeasure } from './morph-text-frame';
import { buildMorphTextGlyphPlan } from './morph-text-glyph-plan';

/** Every character advances by a per-face fraction of the font size. */
const measure: MorphTextMeasure = (text, font) => {
	const [, size = '0', family = ''] = /([\d.]+)px\s+"?([^",]+)/u.exec(font) ?? [];
	return [...text].length * Number(size) * (/Arial/u.test(family) ? 0.52 : 0.47);
};

const slide = (elements: PptxElement[]): PptxSlide => ({ id: 's', elements }) as PptxSlide;

function textBox(
	id: string,
	style: { fontFamily: string; fontSize: number },
	box: Partial<PptxElement>,
	text = 'Lorem ipsum dolor',
): PptxElement {
	return {
		id,
		type: 'text',
		x: 107,
		y: 160,
		width: 466,
		height: 80,
		text,
		textSegments: [{ text, style: { ...style, color: '#FFFFFF' } }],
		textStyle: { ...style, textWrap: 'none' },
		shapeStyle: { fillMode: 'none' },
		...box,
	} as PptxElement;
}

interface Stop {
	percent: number;
	x: number;
	y: number;
	rotation: number;
	scale: number;
	ax: number;
	ay: number;
	opacity: number;
}

/** The stops of one copy's keyframes, read back from the CSS. */
function stopsOf(css: string, name: string): Stop[] {
	const block = css.split(`@keyframes ${name} {`)[1]?.split('\n}')[0] ?? '';
	return [...block.matchAll(/([\d.]+)% \{([^}]*)\}/gu)].map((match) => {
		const body = match[2] ?? '';
		const t =
			/translate\(([-\d.]+)px, ([-\d.]+)px\) rotate\(([-\d.]+)deg\) scale\(([-\d.]+)\) translate\(([-\d.]+)px, ([-\d.]+)px\)/u.exec(
				body,
			) ?? [];
		return {
			percent: Number(match[1]),
			x: Number(t[1]),
			y: Number(t[2]),
			rotation: Number(t[3]),
			scale: Number(t[4]),
			ax: Number(t[5]),
			ay: Number(t[6]),
			opacity: Number(/opacity: ([-\d.]+)/u.exec(body)?.[1]),
		};
	});
}

/**
 * Where a copy's stop sends the glyph's own point: the element's centre plus
 * the stop's translation, because `translate(centre - anchor)` first lays the
 * anchor on the centre.
 */
function glyphAt(stop: Stop, element: PptxElement) {
	return {
		x: element.x + element.width / 2 + stop.x,
		y: element.y + element.height / 2 + stop.y,
	};
}

// PowerPoint 16 (`CreateVideo`, 60 fps, 720p): a line turning while its type
// changes is not turned as one piece. Each glyph's middle moves straight from
// where it stands on one slide to where it stands on the other, so the line
// follows the chord between its two layouts and its length dips below the
// arc's.
describe('a text box that turns while its type changes', () => {
	const from = textBox('a', { fontFamily: 'Calibri', fontSize: 42.67 }, {});
	const to = textBox(
		'b',
		{ fontFamily: 'Arial', fontSize: 53.33 },
		{ x: 107, y: 160, rotation: 60 },
	);
	const plan = buildMorphTextGlyphPlan(
		[{ fromElement: from, toElement: to }],
		new Set(['a']),
		2000,
		measure,
	);
	const css = plan.keyframes.join('\n');
	const outgoing = (index: number) =>
		stopsOf(css, `pptx-morph-glyph-${`a-morph-glyph-${index}`.replace(/[^a-zA-Z0-9]/gu, '')}`);
	const incoming = (index: number) =>
		stopsOf(css, `pptx-morph-glyph-${`b-morph-glyph-${index}`.replace(/[^a-zA-Z0-9]/gu, '')}`);

	it('carries one copy of each half per visible glyph', () => {
		expect(plan.pairs).toHaveLength(1);
		expect(plan.groups).toHaveLength(15);
		expect(plan.groups[0]?.outgoing.id).toBe('a-morph-glyph-0');
		expect(plan.groups[0]?.incoming.id).toBe('b-morph-glyph-0');
		expect(plan.sources.get('a-morph-glyph-3')).toBe('a');
		expect(plan.sources.get('b-morph-glyph-3')).toBe('b');
	});

	it('plays sampled frames on the morph curve, linearly', () => {
		const stops = outgoing(0);
		expect(stops).toHaveLength(25);
		expect(stops[0]?.percent).toBe(0);
		expect(stops.at(-1)?.percent).toBe(100);
		expect(plan.outgoingAnimations.get('a-morph-glyph-0')).toMatch(/2000ms linear forwards$/u);
		// The dissolve rides the same progress as the travel.
		expect(stops[0]?.opacity).toBe(1);
		expect(stops.at(-1)?.opacity).toBe(0);
		expect(incoming(0)[0]?.opacity).toBe(0);
		expect(incoming(0).at(-1)?.opacity).toBe(1);
	});

	it('starts and ends each copy on its own geometry', () => {
		const first = outgoing(0);
		const last = incoming(5);
		expect(first[0]?.rotation).toBe(0);
		expect(first[0]?.scale).toBe(1);
		expect(last.at(-1)?.rotation).toBe(60);
		expect(last.at(-1)?.scale).toBe(1);
	});

	it('moves each glyph straight, and the line along the chord of its two layouts', () => {
		const ends = [0, 14];
		const start = ends.map((index) => glyphAt(outgoing(index)[0]!, from));
		const end = ends.map((index) => glyphAt(incoming(index).at(-1)!, to));
		const length = (a: { x: number; y: number }, b: { x: number; y: number }) =>
			Math.hypot(a.x - b.x, a.y - b.y);
		const stopsA = ends.map((index) => outgoing(index));
		for (const at of [6, 12, 18]) {
			const progress = 1 - (stopsA[0]?.[at]?.opacity ?? 0);
			const a = ends.map((_, k) => glyphAt(stopsA[k]![at]!, from));
			for (const k of [0, 1]) {
				expect(a[k]!.x).toBeCloseTo(start[k]!.x + (end[k]!.x - start[k]!.x) * progress, 0);
				expect(a[k]!.y).toBeCloseTo(start[k]!.y + (end[k]!.y - start[k]!.y) * progress, 0);
			}
			const chord = length(a[0]!, a[1]!);
			const arc =
				length(start[0]!, start[1]!) * (1 - progress) + length(end[0]!, end[1]!) * progress;
			// A rigid turn would hold the straight interpolation of the two lengths.
			expect(chord).toBeLessThan(arc - 5);
		}
	});

	it('turns each glyph about itself on the progress, and sizes it on the progress', () => {
		const stops = outgoing(4);
		const mid = stops[12]!;
		const progress = 1 - mid.opacity;
		expect(mid.rotation).toBeCloseTo(60 * progress, 1);
		expect(mid.scale).toBeGreaterThan(1);
		expect(mid.scale).toBeLessThan(stops.at(-1)!.scale);
		// The incoming copy is the outgoing one at the other face's scale.
		expect(incoming(4)[12]!.scale).toBeCloseTo(mid.scale / stops.at(-1)!.scale, 3);
	});

	it('is what the whole plan paints for the pair', () => {
		const whole = buildMorphTransitionPlan(slide([from]), slide([to]), 2000, 'object', {
			measureText: measure,
		})!;
		expect(whole.keyframesCss).toContain('pptx-morph-glyph-amorphglyph0');
		expect(whole.keyframesCss).toContain('pptx-morph-glyph-bmorphglyph14');
	});

	it('leaves a pair with nothing to restyle, or nothing turning, to the box', () => {
		const alone = (a: PptxElement, b: PptxElement) =>
			buildMorphTextGlyphPlan([{ fromElement: a, toElement: b }], new Set(['a']), 2000, measure);
		const same = textBox('b', { fontFamily: 'Calibri', fontSize: 42.67 }, { rotation: 60 });
		const still = textBox('b', { fontFamily: 'Arial', fontSize: 53.33 }, { x: 200 });
		expect(alone(from, same).pairs).toHaveLength(0);
		expect(alone(from, still).pairs).toHaveLength(0);
	});

	it('leaves several lines to the box', () => {
		const two = (id: string, style: { fontFamily: string; fontSize: number }, rotation: number) =>
			textBox(id, style, { rotation }, 'Lorem ipsum\nsit amet');
		const result = buildMorphTextGlyphPlan(
			[
				{
					fromElement: two('a', { fontFamily: 'Calibri', fontSize: 42.67 }, 0),
					toElement: two('b', { fontFamily: 'Arial', fontSize: 53.33 }, 60),
				},
			],
			new Set(['a']),
			2000,
			measure,
		);
		expect(result.pairs).toHaveLength(0);
	});
});

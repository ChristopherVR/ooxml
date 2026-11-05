import type { PptxElement, PptxSlide } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import { buildMorphTransitionPlan } from './morph-plan';
import type { MorphTextMeasure, MorphTextTransform } from './morph-text-frame';
import {
	morphTextGhostEnd,
	morphTextIncomingStart,
	resolveMorphTextAlignment,
} from './morph-text-frame';

/**
 * A stand-in for the canvas: every character advances by a per-face fraction
 * of the font size, so Calibri Light is narrower than Calibri, and Arial wider.
 */
const measure: MorphTextMeasure = (text, font) => {
	const [, size = '0', family = ''] = /([\d.]+)px\s+"?([^",]+)/u.exec(font) ?? [];
	const face = /Light/u.test(family) ? 0.45 : /Arial/u.test(family) ? 0.52 : 0.47;
	return [...text].length * Number(size) * face;
};

const noMeasure: MorphTextMeasure = () => undefined;

function textBox(
	id: string,
	paragraphs: string[],
	style: { fontFamily: string; fontSize: number },
	geometry: Partial<PptxElement> & { align?: 'left' | 'center' | 'right' },
): PptxElement {
	const { align, ...box } = geometry;
	const segments = paragraphs.flatMap((text, index) => [
		...(index > 0 ? [{ text: '\n' }] : []),
		{ text, style: { ...style, color: '#FFFFFF' } },
	]);
	return {
		id,
		type: 'text',
		x: 0,
		y: 0,
		width: 1173,
		height: 80,
		text: paragraphs.join('\n'),
		textSegments: segments,
		textStyle: {
			...style,
			bodyInsetLeft: 0,
			bodyInsetRight: 0,
			bodyInsetTop: 0,
			bodyInsetBottom: 0,
			...(align ? { align } : {}),
		},
		shapeStyle: { fillMode: 'none' },
		...box,
	} as PptxElement;
}

/** Where `transform` (about the element's centre) sends the slide point `x, y`. */
function apply(t: MorphTextTransform, element: PptxElement, x: number, y: number) {
	const cx = element.x + element.width / 2;
	const cy = element.y + element.height / 2;
	return { x: cx + t.dx + t.sx * (x - cx), y: cy + t.dy + t.sy * (y - cy) };
}

const slide = (elements: PptxElement[]): PptxSlide => ({ id: 's', elements }) as PptxSlide;

// Issue #141, the corporate deck's slides 6 -> 7: "Object 6" is one long line
// of 54 pt text in a 3314 px box that starts wholly off the RIGHT edge and
// flies far left, changing Calibri Light -> Calibri on the way. PowerPoint
// shows one copy at every frame; the box interpolation drew two copies whose
// glyphs drifted apart along the line.
describe('carried text lines (issue #141, slides 6 -> 7)', () => {
	const line = 'Lorem ipsu LoremLore LoremLo ipsumip dolordo Lor ips dol sits Lo ip dolordolor';
	const from = textBox(
		'o6-6',
		[line],
		{ fontFamily: 'Calibri Light', fontSize: 72 },
		{ x: 1339, y: 273, width: 3314, height: 87 },
	);
	const to = textBox(
		'o6-7',
		[line],
		{ fontFamily: 'Calibri', fontSize: 72 },
		{ x: -2252, y: 273, width: 3314, height: 87 },
	);

	it('maps the outgoing line onto the incoming line by its width ratio', () => {
		const alignment = resolveMorphTextAlignment(from, to, measure)!;
		expect(alignment).toBeDefined();
		expect(alignment.scaleX).toBeCloseTo(0.47 / 0.45, 6);
		expect(alignment.scaleY).toBe(1);
		expect(alignment.from).toEqual({ x: 1339, y: 273 });
		expect(alignment.to).toEqual({ x: -2252, y: 273 });
	});

	it('starts the incoming line on the outgoing one and lands the ghost on the incoming one', () => {
		const alignment = resolveMorphTextAlignment(from, to, measure)!;
		const widthFrom = [...line].length * 72 * 0.45;
		const widthTo = [...line].length * 72 * 0.47;
		const start = morphTextIncomingStart(alignment, to);
		expect(apply(start, to, -2252, 273).x).toBeCloseTo(1339, 6);
		expect(apply(start, to, -2252 + widthTo, 273).x).toBeCloseTo(1339 + widthFrom, 6);
		const end = morphTextGhostEnd(alignment, from);
		expect(apply(end, from, 1339, 273).x).toBeCloseTo(-2252, 6);
		expect(apply(end, from, 1339 + widthFrom, 360).x).toBeCloseTo(-2252 + widthTo, 6);
		expect(apply(end, from, 1339, 360).y).toBeCloseTo(360, 6);
	});

	it('scales both halves of the plan by the line ratio, not the box ratio', () => {
		const plan = buildMorphTransitionPlan(slide([from]), slide([to]), 2000, 'object', {
			measureText: measure,
		})!;
		const ratio = 0.47 / 0.45;
		const incoming = plan.keyframesCss.split('@keyframes pptx-morph-0-')[1] ?? '';
		expect(incoming).toContain(`scale(${1 / ratio}, 1)`);
		const ghost = plan.keyframesCss.split('@keyframes pptx-morph-ghost-')[1] ?? '';
		expect(ghost).toContain(`scale(${ratio}, 1)`);
	});

	it('dissolves the two halves additively when nothing is stacked over them', () => {
		const plan = buildMorphTransitionPlan(slide([from]), slide([to]), 2000, 'object', {
			measureText: measure,
		})!;
		expect(plan.crossfadeGroups.map((group) => group.incoming.id)).toEqual(['o6-7']);
		// It travels, so its journey stays on the element.
		expect(plan.crossfadeGroups[0]?.incomingAnimation).toBeUndefined();
	});

	it('keeps the stacking when a shape above it shares its path', () => {
		const cover = {
			id: 'cover',
			type: 'shape',
			x: 0,
			y: 250,
			width: 600,
			height: 150,
			shapeStyle: { fillColor: '#336699' },
		} as PptxElement;
		const plan = buildMorphTransitionPlan(
			slide([from, cover]),
			slide([to, { ...cover, id: 'cover-7' }]),
			2000,
			'object',
			{ measureText: measure },
		)!;
		expect(plan.crossfadeGroups).toHaveLength(0);
	});

	it('leaves word and character morphs to their tokens', () => {
		const plan = buildMorphTransitionPlan(slide([from]), slide([to]), 2000, 'word', {
			measureText: measure,
		})!;
		expect(plan.keyframesCss).not.toContain(`scale(${0.47 / 0.45}, 1)`);
	});
});

// Issue #141, slides 17 -> 18 and back: a size-only change flown in from wholly
// off the top or left edge stays one copy, painted in the incoming layout, even
// when the text can be measured.
describe('off-stage size-only flights keep one copy', () => {
	const off = textBox(
		'off',
		['Heading', 'Body'],
		{ fontFamily: 'Arial', fontSize: 14 },
		{ x: -500, width: 200, height: 200 },
	);
	const on = textBox(
		'on',
		['Heading', 'Body'],
		{ fontFamily: 'Arial', fontSize: 18 },
		{ x: 300, width: 200, height: 210 },
	);

	it('flies forward and back without a ghost dissolve', () => {
		for (const [a, b] of [
			[off, on],
			[{ ...on, id: 'on2', x: -500 } as PptxElement, { ...off, id: 'off2', x: 300 } as PptxElement],
		] as const) {
			const plan = buildMorphTransitionPlan(slide([a]), slide([b]), 1000, 'object', {
				measureText: measure,
			})!;
			expect(plan.outgoingElements.filter((element) => element.id === a.id)).toHaveLength(0);
			expect(plan.incomingAnimations.get(b.id)).not.toContain('-fade');
			expect(plan.crossfadeGroups).toHaveLength(0);
		}
	});
});

// The deck's "L" columns: twelve one-letter paragraphs that change face while
// the column climbs.
describe('carried text columns', () => {
	const ls = Array.from({ length: 12 }, () => 'L');
	const from = textBox(
		'l-6',
		ls,
		{ fontFamily: 'Calibri Light', fontSize: 72 },
		{ x: 284, y: -916, width: 51, height: 1067 },
	);
	const to = textBox(
		'l-7',
		ls,
		{ fontFamily: 'Calibri', fontSize: 72 },
		{ x: 284, y: -158, width: 51, height: 1067 },
	);

	it('carries every line together', () => {
		const alignment = resolveMorphTextAlignment(from, to, measure)!;
		expect(alignment.scaleX).toBeCloseTo(0.47 / 0.45, 6);
		expect(alignment.scaleY).toBe(1);
		expect(alignment.to.y - alignment.from.y).toBe(758);
	});
});

// Synthetic PowerPoint 16 decks (CreateVideo, 720p): a 24 -> 48 pt change
// scales one copy uniformly about the line's start; a centred line grows
// about its centre (206..1075 px -> 169..1110 px, centre held at 640 px).
describe('carried text geometry', () => {
	it('scales a size change uniformly about the line start', () => {
		const from = textBox(
			'a',
			['Lorem ipsum dolor sit'],
			{ fontFamily: 'Calibri', fontSize: 32 },
			{ x: 53, y: 267, width: 1173, height: 40 },
		);
		const to = textBox(
			'b',
			['Lorem ipsum dolor sit'],
			{ fontFamily: 'Calibri', fontSize: 64 },
			{ x: 53, y: 267, width: 1173, height: 80 },
		);
		const alignment = resolveMorphTextAlignment(from, to, measure)!;
		expect(alignment.scaleX).toBeCloseTo(2, 6);
		expect(alignment.scaleY).toBeCloseTo(2, 6);
		expect(alignment.from).toEqual(alignment.to);
	});

	it('anchors a centred line at its centre', () => {
		const geometry = { x: 53, y: 200, width: 1173, height: 80, align: 'center' as const };
		const text = ['Lorem ipsum dolor sit amet consectetur adipiscing'];
		const from = textBox('a', text, { fontFamily: 'Calibri', fontSize: 42.67 }, geometry);
		const to = textBox(
			'b',
			text,
			{ fontFamily: 'Arial', fontSize: 42.67 },
			{ ...geometry, y: 360 },
		);
		const alignment = resolveMorphTextAlignment(from, to, measure)!;
		expect(alignment.from.x).toBeCloseTo(53 + 1173 / 2, 6);
		expect(alignment.to.x).toBeCloseTo(53 + 1173 / 2, 6);
		const end = morphTextGhostEnd(alignment, from);
		expect(apply(end, from, 53 + 1173 / 2, 200).x).toBeCloseTo(53 + 1173 / 2, 6);
	});

	it('moves unchanged text rigidly without measuring, whatever the box does', () => {
		const from = textBox(
			'a',
			['Title'],
			{ fontFamily: 'Calibri', fontSize: 40 },
			{ x: 100, y: 100, width: 400, height: 60 },
		);
		const to = textBox(
			'b',
			['Title'],
			{ fontFamily: 'Calibri', fontSize: 40 },
			{ x: 300, y: 200, width: 400, height: 120 },
		);
		const alignment = resolveMorphTextAlignment(from, to, noMeasure)!;
		expect(alignment).toMatchObject({ scaleX: 1, scaleY: 1 });
		const start = morphTextIncomingStart(alignment, to);
		expect(start.sx).toBe(1);
		expect(start.sy).toBe(1);
		expect(apply(start, to, 300, 200)).toEqual({ x: 100, y: 100 });
	});

	it('falls back to the box when a layout may wrap differently', () => {
		const long = ['Lorem ipsum dolor sit amet consectetur adipiscing elit sed do eiusmod'];
		const from = textBox(
			'a',
			long,
			{ fontFamily: 'Calibri', fontSize: 28 },
			{ x: 0, y: 0, width: 667, height: 80 },
		);
		const to = textBox(
			'b',
			long,
			{ fontFamily: 'Calibri', fontSize: 36 },
			{ x: 0, y: 0, width: 667, height: 160 },
		);
		expect(resolveMorphTextAlignment(from, to, measure)).toBeUndefined();
	});

	it('falls back without a measurer, for another face, or for a painted body', () => {
		const from = textBox(
			'a',
			['Title'],
			{ fontFamily: 'Calibri Light', fontSize: 40 },
			{ x: 0, y: 0, width: 400, height: 60 },
		);
		const to = textBox(
			'b',
			['Title'],
			{ fontFamily: 'Calibri', fontSize: 40 },
			{ x: 300, y: 0, width: 400, height: 60 },
		);
		expect(resolveMorphTextAlignment(from, to, noMeasure)).toBeUndefined();
		expect(resolveMorphTextAlignment(from, { ...to, text: 'Other' } as PptxElement, measure)).toBe(
			undefined,
		);
		const filled = { ...to, shapeStyle: { fillColor: '#ff0000' } } as PptxElement;
		expect(resolveMorphTextAlignment(from, filled, measure)).toBeUndefined();
		const rotated = { ...to, rotation: 30 } as PptxElement;
		expect(resolveMorphTextAlignment(from, rotated, measure)).toBeUndefined();
	});
});

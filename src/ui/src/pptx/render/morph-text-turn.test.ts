import type { PptxElement, PptxSlide } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import { buildMorphTransitionPlan } from './morph-plan';
import type { MorphTextMeasure, MorphTextTransform } from './morph-text-frame';
import {
	morphTextGhostEnd,
	morphTextGhostStart,
	morphTextIncomingEnd,
	morphTextIncomingStart,
	resolveMorphTextAlignment,
} from './morph-text-frame';

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
): PptxElement {
	return {
		id,
		type: 'text',
		x: 107,
		y: 160,
		width: 466,
		height: 80,
		text: 'Lorem ipsum dolor',
		textSegments: [{ text: 'Lorem ipsum dolor', style: { ...style, color: '#FFFFFF' } }],
		textStyle: { ...style, textWrap: 'none' },
		shapeStyle: { fillMode: 'none' },
		...box,
	} as PptxElement;
}

/**
 * Where the CSS frame `translate(d) rotate(r) scale(s) translate(-tail)`
 * (about the element's centre) sends the element's unrotated slide point.
 */
function place(
	t: MorphTextTransform,
	rotation: number,
	element: PptxElement,
	x: number,
	y: number,
) {
	const cx = element.x + element.width / 2;
	const cy = element.y + element.height / 2;
	const vx = t.sx * (x - cx - (t.tail?.x ?? 0));
	const vy = t.sy * (y - cy - (t.tail?.y ?? 0));
	const r = (rotation * Math.PI) / 180;
	return {
		x: cx + t.dx + vx * Math.cos(r) - vy * Math.sin(r),
		y: cy + t.dy + vx * Math.sin(r) + vy * Math.cos(r),
	};
}

/** CSS's per-function interpolation of two frames sharing one function list. */
function mix(a: MorphTextTransform, b: MorphTextTransform, p: number): MorphTextTransform {
	const at = (u: number, v: number) => u + (v - u) * p;
	return {
		dx: at(a.dx, b.dx),
		dy: at(a.dy, b.dy),
		sx: at(a.sx, b.sx),
		sy: at(a.sy, b.sy),
		...(a.tail ? { tail: a.tail } : {}),
	};
}

// PowerPoint 16 (`CreateVideo`, 720p): a 32 pt line in a box turning 0 -> 40
// degrees while it moves and restyles keeps one glyph-aligned copy, and the
// middle of its text travels in a straight line at the rotation's progress.
describe('a text box that turns while it morphs', () => {
	const from = textBox('a', { fontFamily: 'Calibri', fontSize: 42.67 }, {});
	const to = textBox(
		'b',
		{ fontFamily: 'Arial', fontSize: 42.67 },
		{ x: 560, y: 427, rotation: 40 },
	);
	const alignment = resolveMorphTextAlignment(from, to, measure)!;

	it('turns about the middle of its text', () => {
		const width = 17 * 42.67 * 0.47;
		expect(alignment.turn?.from).toBe(0);
		expect(alignment.turn?.to).toBe(40);
		expect(alignment.turn?.pivots.from.x).toBeCloseTo(alignment.from.x + width / 2, 6);
	});

	it('keeps both halves glyph over glyph on every frame', () => {
		const ghost = [morphTextGhostStart(alignment, from), morphTextGhostEnd(alignment, from)];
		const incoming = [morphTextIncomingStart(alignment, to), morphTextIncomingEnd(alignment, to)];
		const ratio = 0.52 / 0.47;
		for (const p of [0, 0.25, 0.5, 0.75, 1]) {
			const rotation = 40 * p;
			const g = mix(ghost[0]!, ghost[1]!, p);
			const i = mix(incoming[0]!, incoming[1]!, p);
			// The end of the line in each half.
			const gEnd = place(g, rotation, from, alignment.from.x + 300, alignment.from.y);
			const iEnd = place(i, rotation, to, alignment.to.x + 300 * ratio, alignment.to.y);
			expect(iEnd.x).toBeCloseTo(gEnd.x, 6);
			expect(iEnd.y).toBeCloseTo(gEnd.y, 6);
		}
		// The incoming half ends at its own geometry: `rotate(40deg)` alone.
		const own = place(incoming[1]!, 40, to, to.x, to.y);
		const plain = place({ dx: 0, dy: 0, sx: 1, sy: 1 }, 40, to, to.x, to.y);
		expect(own.x).toBeCloseTo(plain.x, 6);
		expect(own.y).toBeCloseTo(plain.y, 6);
	});

	it('moves the middle of its text in a straight line', () => {
		const ghost = [morphTextGhostStart(alignment, from), morphTextGhostEnd(alignment, from)];
		const pivot = alignment.turn!.pivots.from;
		const at = (p: number) => place(mix(ghost[0]!, ghost[1]!, p), 40 * p, from, pivot.x, pivot.y);
		const [start, middle, end] = [at(0), at(0.5), at(1)];
		expect(middle.x).toBeCloseTo((start.x + end.x) / 2, 6);
		expect(middle.y).toBeCloseTo((start.y + end.y) / 2, 6);
	});

	it('never stretches unchanged text with its box', () => {
		const wide = textBox(
			'c',
			{ fontFamily: 'Calibri', fontSize: 42.67 },
			{ x: 520, y: 537, width: 693, rotation: 40 },
		);
		const plan = buildMorphTransitionPlan(slide([from]), slide([wide]), 2000, 'object', {
			measureText: measure,
		})!;
		const incoming = plan.keyframesCss.split('@keyframes pptx-morph-0-')[1] ?? '';
		expect(incoming).toContain('scale(1, 1)');
		expect(incoming).not.toContain('scale(0.672');
		// Both frames close on the pivot, so the list interpolates function by function.
		expect(incoming.match(/scaleY\(1\) translate\(/gu)).toHaveLength(2);
	});
});

// PowerPoint 16: a text box whose horizontal flip changes between the slides
// moves and restyles like an unflipped one; its text never mirrors.
describe('a text box whose flip changes', () => {
	const from = textBox('a', { fontFamily: 'Calibri', fontSize: 42.67 }, {});
	const to = textBox(
		'b',
		{ fontFamily: 'Arial', fontSize: 42.67 },
		{ x: 560, y: 427, flipHorizontal: true },
	);

	it('carries the text and holds each half at its own flips', () => {
		expect(resolveMorphTextAlignment(from, to, measure)).toBeDefined();
		const plan = buildMorphTransitionPlan(slide([from]), slide([to]), 2000, 'object', {
			measureText: measure,
		})!;
		const block = (name: string) =>
			plan.keyframesCss.split(`@keyframes ${name}`)[1]?.split('@keyframes')[0] ?? '';
		// Both frames of the incoming half state its flip; neither of the ghost's does.
		expect(block('pptx-morph-0-').match(/scaleX\(-1\)/gu)).toHaveLength(2);
		expect(block('pptx-morph-ghost-')).not.toContain('scaleX(-1)');
	});

	it('carries a vertical flip as a half turn of the text (see morph-text-flip)', () => {
		const upside = { ...to, flipHorizontal: false, flipVertical: true } as PptxElement;
		expect(resolveMorphTextAlignment(from, upside, measure)?.turn?.to).toBe(180);
	});

	it('refuses a flip whose insets would move the text', () => {
		const lopsided = {
			...to,
			textStyle: { ...(to as { textStyle: object }).textStyle, bodyInsetLeft: 40 },
		} as PptxElement;
		expect(resolveMorphTextAlignment(from, lopsided, measure)).toBeUndefined();
	});
});

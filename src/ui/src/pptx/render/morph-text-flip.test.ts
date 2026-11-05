import type { PptxElement, PptxSlide } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import { buildMorphTransitionPlan } from './morph-plan';
import { morphCarriedTextRotations, morphTextUpright } from './morph-text-flip';
import type { MorphTextMeasure, MorphTextTransform } from './morph-text-frame';
import {
	morphTextGhostEnd,
	morphTextGhostStart,
	morphTextIncomingEnd,
	morphTextIncomingStart,
	resolveMorphTextAlignment,
} from './morph-text-frame';
import { morphTextTail } from './morph-text-transform';

const measure: MorphTextMeasure = (text, font) => {
	const [, size = '0'] = /([\d.]+)px/u.exec(font) ?? [];
	return [...text].length * Number(size) * 0.47;
};

const slide = (elements: PptxElement[]): PptxSlide => ({ id: 's', elements }) as PptxSlide;

function textBox(id: string, box: Partial<PptxElement>): PptxElement {
	const style = { fontFamily: 'Calibri', fontSize: 42.67 };
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

type Point = { x: number; y: number };

const turn = (p: Point, degrees: number): Point => {
	const r = (degrees * Math.PI) / 180;
	return { x: p.x * Math.cos(r) - p.y * Math.sin(r), y: p.x * Math.sin(r) + p.y * Math.cos(r) };
};

/**
 * Where a copy painting `element` (its own flips, its text compensated as the
 * static renderer does) puts its text point `p` (element-local, from the
 * centre) under the CSS frame `translate(d) rotate(r) scale(s) flips tail`.
 */
function painted(t: MorphTextTransform, rotation: number, element: PptxElement, p: Point): Point {
	const [, ux = '0', uy = '0'] =
		/translate\(([-\d.]+)px, ([-\d.]+)px\)/u.exec(morphTextTail(t, element)) ?? [];
	const mirrored = Boolean(element.flipHorizontal) !== Boolean(element.flipVertical);
	const compensated = { x: mirrored ? -p.x : p.x, y: p.y };
	const shifted = { x: compensated.x + Number(ux), y: compensated.y + Number(uy) };
	const flipped = {
		x: element.flipHorizontal ? -shifted.x : shifted.x,
		y: element.flipVertical ? -shifted.y : shifted.y,
	};
	const v = turn({ x: t.sx * flipped.x, y: t.sy * flipped.y }, rotation);
	return {
		x: element.x + element.width / 2 + t.dx + v.x,
		y: element.y + element.height / 2 + t.dy + v.y,
	};
}

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

describe('morphTextUpright', () => {
	it('turns a vertical flip into a half turn and drops a horizontal one', () => {
		const box = textBox('a', { rotation: 30 });
		expect(morphTextUpright(box)).toBe(box);
		expect(morphTextUpright({ ...box, flipVertical: true } as PptxElement)).toMatchObject({
			rotation: 210,
			flipHorizontal: false,
			flipVertical: false,
		});
		expect(morphTextUpright({ ...box, flipHorizontal: true } as PptxElement)?.rotation).toBe(30);
		const both = { ...box, rotation: 240, flipHorizontal: true, flipVertical: true };
		expect(morphTextUpright(both as PptxElement)?.rotation).toBe(60);
	});

	it('refuses a flip whose text rectangle sits off centre', () => {
		const lopsided = textBox('a', { flipVertical: true });
		(lopsided as { textStyle: { bodyInsetLeft?: number } }).textStyle.bodyInsetLeft = 40;
		expect(morphTextUpright(lopsided)).toBeUndefined();
	});
});

// PowerPoint 16 (`CreateVideo`, 60 fps, 720p): a text box gaining `flipV`
// while it moves turns its text 0 -> -180 degrees (anti-clockwise) about the
// middle of its text, which stays 318 px long on every frame.
describe('a text box whose vertical flip changes', () => {
	const from = textBox('a', {});
	const to = textBox('b', { x: 560, y: 427, flipVertical: true });
	const alignment = resolveMorphTextAlignment(from, to, measure)!;
	const rotations = morphCarriedTextRotations(from, to);

	it('carries the text as a half turn, anti-clockwise', () => {
		expect(alignment.turn?.from).toBe(0);
		expect(alignment.turn?.to).toBe(180);
		expect(rotations.ghost).toEqual([0, -180]);
		// The incoming half paints its text at its rotation + 180.
		expect(rotations.incoming).toEqual([180, 0]);
		const back = morphCarriedTextRotations(to, from);
		// Losing the flip turns from 180 the PowerPoint half-turn way (clockwise from 180).
		expect(back.ghost[1] + 180 - (back.ghost[0] + 180)).toBe(180);
	});

	it('keeps both halves glyph over glyph through the turn', () => {
		const ghost = [morphTextGhostStart(alignment, from), morphTextGhostEnd(alignment, from)];
		const incoming = [morphTextIncomingStart(alignment, to), morphTextIncomingEnd(alignment, to)];
		const lineEnd = { x: 300 - from.width / 2 + 9.6, y: -20 };
		for (const p of [0, 0.25, 0.5, 0.75, 1]) {
			const g = painted(
				mix(ghost[0]!, ghost[1]!, p),
				rotations.ghost[0] + (rotations.ghost[1] - rotations.ghost[0]) * p,
				from,
				lineEnd,
			);
			const i = painted(
				mix(incoming[0]!, incoming[1]!, p),
				rotations.incoming[0] + (rotations.incoming[1] - rotations.incoming[0]) * p,
				to,
				lineEnd,
			);
			expect(i.x).toBeCloseTo(g.x, 2);
			expect(i.y).toBeCloseTo(g.y, 2);
		}
		// The incoming half ends where the static, flipped box paints it.
		const own = painted(incoming[1]!, 0, to, lineEnd);
		const still = painted({ dx: 0, dy: 0, sx: 1, sy: 1 }, 0, to, lineEnd);
		expect(own.x).toBeCloseTo(still.x, 2);
		expect(own.y).toBeCloseTo(still.y, 2);
	});

	it('states each half at its own flips in the plan', () => {
		const plan = buildMorphTransitionPlan(slide([from]), slide([to]), 2000, 'object', {
			measureText: measure,
		})!;
		const block = (name: string) =>
			plan.keyframesCss.split(`@keyframes ${name}`)[1]?.split('@keyframes')[0] ?? '';
		const incoming = block('pptx-morph-0-');
		expect(incoming.match(/scaleY\(-1\)/gu)).toHaveLength(2);
		expect(incoming).toContain('rotate(180deg)');
		const ghost = block('pptx-morph-ghost-');
		if (ghost) {
			expect(ghost).toContain('rotate(-180deg)');
		}
	});
});

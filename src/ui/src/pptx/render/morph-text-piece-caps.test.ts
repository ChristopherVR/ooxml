import type { PptxElement } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import type { MorphFontMetrics } from './morph-text-baseline';
import type { MorphTextMeasure } from './morph-text-frame';
import { resolveMorphTextAlignment } from './morph-text-frame';
import type { MorphTextInkMeasure } from './morph-text-ink';
import { resolveMorphTextPieces } from './morph-text-pieces';

/** Arial is wider than Calibri. */
const measure: MorphTextMeasure = (text, font) => {
	const [, size = '0', family = ''] = /([\d.]+)px\s+"?([^",]+)/u.exec(font) ?? [];
	return [...text].length * Number(size) * (/Arial/u.test(family) ? 0.52 : 0.47);
};

const first = (font: string) => /px\s+"?([^",]+)/u.exec(font)?.[1] ?? '';
const capRatio = (font: string) => (first(font) === 'Arial' ? 0.72 : 0.6);
const sizeOf = (font: string) => Number(/([\d.]+)px/u.exec(font)?.[1] ?? 0);

/** Arial's capitals are taller than Calibri's; both sit on the same ascent. */
const metrics: MorphFontMetrics = (font) => ({
	ascent: 0.9 * sizeOf(font),
	descent: 0.2 * sizeOf(font),
	capHeight: capRatio(font) * sizeOf(font),
});

const ink: MorphTextInkMeasure = (text, font) => ({
	left: 0,
	right: [...text].length * sizeOf(font) * 0.5,
	ascent: capRatio(font) * sizeOf(font),
	descent: 0,
	fontAscent: 0.9 * sizeOf(font),
	fontDescent: 0.2 * sizeOf(font),
});

type Run = { text: string; fontFamily: string };

/** A text box of paragraphs of runs, 32 px, no insets, wide enough not to wrap. */
function textBox(id: string, paragraphs: Run[][], box: Partial<PptxElement> = {}): PptxElement {
	const segments = paragraphs.flatMap((runs, index) => [
		...(index > 0 ? [{ text: '\n' }] : []),
		...runs.map((run) => ({
			text: run.text,
			style: { fontFamily: run.fontFamily, fontSize: 32, color: '#FFFFFF' },
		})),
	]);
	return {
		id,
		type: 'text',
		x: 80,
		y: 100,
		width: 800,
		height: 200,
		text: paragraphs.map((runs) => runs.map((run) => run.text).join('')).join('\n'),
		textSegments: segments,
		textStyle: {
			fontFamily: paragraphs[0]?.[0]?.fontFamily ?? 'Calibri',
			fontSize: 32,
			bodyInsetLeft: 0,
			bodyInsetRight: 0,
			bodyInsetTop: 0,
			bodyInsetBottom: 0,
		},
		shapeStyle: { fillMode: 'none' },
		...box,
	} as PptxElement;
}

const lines = ['Lorem ipsum dolor', 'Sit amet elit'];
const pair = (fromFace: Run['fontFamily'], toFace: Run['fontFamily']) => [
	textBox(
		'a',
		lines.map((text) => [{ text, fontFamily: fromFace }]),
	),
	textBox(
		'b',
		lines.map((text) => [{ text, fontFamily: toFace }]),
		{ x: 380, y: 300 },
	),
];

// PowerPoint 16 (`CreateVideo`, 32 pt, two lines Calibri -> Arial, the same
// breaks): each line's "L" grows 27 -> 31 px while the baselines stay 51-52 px
// apart, so every line is its own piece, scaled about its own baseline.
describe('a change of face over several lines', () => {
	const [from, to] = pair('Calibri', 'Arial') as [PptxElement, PptxElement];

	it('is not carried by one box mapping', () => {
		expect(resolveMorphTextAlignment(from, to, measure, { metrics, ink })).toBeUndefined();
	});

	it('is carried one piece per line, about the line baseline, by the capital ratio', () => {
		const split = resolveMorphTextPieces(from, to, measure, ink, {
			metrics,
			requireCaps: true,
		})!;
		expect(split.pieces).toHaveLength(2);
		for (const [index, piece] of split.pieces.entries()) {
			expect(piece.alignment.scaleY).toBeCloseTo(0.72 / 0.6, 6);
			// The anchor is the line's baseline: 0.9 em in a 1.2 em line box.
			const lineTop = from.y + index * 32 * 1.2;
			expect(piece.alignment.from.y - lineTop).toBeCloseTo(0.9 * 32 + (1.2 * 32 - 1.1 * 32) / 2, 4);
			expect(piece.alignment.to.y - (to.y + index * 32 * 1.2)).toBeCloseTo(
				0.9 * 32 + (1.2 * 32 - 1.1 * 32) / 2,
				4,
			);
		}
		expect(split.pieces[1]!.alignment.from.y).toBeGreaterThan(split.pieces[0]!.alignment.from.y);
	});

	it('keeps the box mapping where nothing can place a baseline', () => {
		expect(resolveMorphTextAlignment(from, to, measure, { metrics })).toBeDefined();
	});
});

// The same face change on one line holding two faces: each face stretches onto
// its own counterpart's capitals (Calibri -> Arial, Georgia -> Times New Roman).
describe('a change of face in several faces on one line', () => {
	const runs = (a: string, b: string): Run[] => [
		{ text: 'Lorem ', fontFamily: a },
		{ text: 'ipsum ', fontFamily: b },
		{ text: 'dolor', fontFamily: a },
	];
	const from = textBox('m1', [runs('Calibri', 'Georgia')]);
	const to = textBox('m2', [runs('Arial', 'Times New Roman')], { x: 300, y: 300 });
	const faces: MorphFontMetrics = (font) => {
		const base = metrics(font)!;
		const size = sizeOf(font);
		return first(font) === 'Georgia'
			? { ...base, capHeight: 0.69 * size }
			: first(font) === 'Times New Roman'
				? { ...base, capHeight: 0.66 * size }
				: base;
	};

	it('flies each run of one face as a piece on its own capital ratio', () => {
		expect(resolveMorphTextAlignment(from, to, measure, { metrics: faces, ink })).toBeUndefined();
		const pieces = resolveMorphTextPieces(from, to, measure, ink, {
			metrics: faces,
			requireCaps: true,
		})!.pieces;
		expect(pieces.map((piece) => piece.alignment.scaleY)).toEqual([
			expect.closeTo(0.72 / 0.6, 6),
			expect.closeTo(0.66 / 0.69, 6),
			expect.closeTo(0.72 / 0.6, 6),
		]);
		// One line: all three share its baseline.
		expect(new Set(pieces.map((piece) => piece.alignment.from.y)).size).toBe(1);
	});
});

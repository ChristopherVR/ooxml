import type { PptxElement, PptxSlide } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import { buildMorphTransitionPlan } from './morph-plan';
import type { MorphTextMeasure, MorphTextTransform } from './morph-text-frame';
import { morphTextGhostEnd, morphTextIncomingStart } from './morph-text-frame';
import { readFrame } from './morph-text-frame-read';
import { layoutMorphTextWords } from './morph-text-lines';
import { morphTextPieceId } from './morph-text-piece-plan';
import { resolveMorphTextPieces } from './morph-text-pieces';

/** A monospace stand-in for the canvas: every character is half the font size wide. */
const measure: MorphTextMeasure = (text, font) => {
	const [, size = '0'] = /([\d.]+)px/u.exec(font) ?? [];
	return [...text].length * Number(size) * 0.5;
};

const TEXT =
	'Lorem ipsum dolor sit amet consectetur adipiscing elit sed do eiusmod tempor incididunt';

function textBox(id: string, fontSize: number, box: Partial<PptxElement> = {}): PptxElement {
	const style = { fontFamily: 'Calibri', fontSize };
	return {
		id,
		type: 'text',
		x: 80,
		y: 133,
		width: 400,
		height: 120,
		text: TEXT,
		textSegments: [{ text: TEXT, style: { ...style, color: '#000000' } }],
		textStyle: {
			...style,
			bodyInsetLeft: 0,
			bodyInsetRight: 0,
			bodyInsetTop: 0,
			bodyInsetBottom: 0,
		},
		shapeStyle: { fillMode: 'none' },
		...box,
	} as PptxElement;
}

function apply(t: MorphTextTransform, element: PptxElement, x: number, y: number) {
	const cx = element.x + element.width / 2;
	const cy = element.y + element.height / 2;
	return { x: cx + t.dx + t.sx * (x - cx), y: cy + t.dy + t.sy * (y - cy) };
}

const slide = (elements: PptxElement[]): PptxSlide => ({ id: 's', elements }) as PptxSlide;

// Modelled on PowerPoint 16's render of one Calibri paragraph in a fixed box,
// 28 pt on the first slide and 36 pt on the second: the two sizes break the
// paragraph differently, and PowerPoint flies the five runs of words that
// share a line in BOTH layouts on their own, each from its old place and size
// to its new one. Here the sizes are 20 and 22 px of a monospace face in a
// 400 px column, which breaks into the same lines.
describe('rewrapped text carried piece by piece', () => {
	const from = textBox('wrap-1', 20);
	const to = textBox('wrap-2', 22);

	it('lays each half out word by word', () => {
		const a = layoutMorphTextWords(readFrame(from)!, measure)!;
		const b = layoutMorphTextWords(readFrame(to)!, measure)!;
		expect(a.lines.map((line) => line.top)).toEqual([0, 24, 48]);
		expect(b.lines.map((line) => line.top)).toEqual(
			[0, 26.4, 52.8].map((v) => expect.closeTo(v, 6)),
		);
		expect(a.words.map((word) => word.line)).toEqual([0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 1, 1, 2]);
		expect(b.words.map((word) => word.line)).toEqual([0, 0, 0, 0, 0, 1, 1, 1, 1, 1, 2, 2, 2]);
		expect(a.words[5]).toMatchObject({ x0: 270, x1: 380 });
	});

	it('splits the pair into the runs of words that share a line in both layouts', () => {
		const pieces = resolveMorphTextPieces(from, to, measure)!.pieces;
		expect(pieces).toHaveLength(5);
		// "consectetur" leaves the end of line 1 for the start of line 2.
		expect(pieces[1]!.alignment).toEqual({
			from: { x: 80 + 270, y: 133 },
			to: { x: 80, y: expect.closeTo(133 + 26.4, 6) },
			scaleX: expect.closeTo(1.1, 6),
			scaleY: expect.closeTo(1.1, 6),
		});
		// "adipiscing elit sed do" stays on line 2 but moves along it.
		expect(pieces[2]!.alignment.from).toEqual({ x: 80, y: 133 + 24 });
		expect(pieces[2]!.alignment.to.x).toBeCloseTo(80 + 132, 6);
		// "incididunt" moves along the last line.
		expect(pieces[4]!.alignment.to.x).toBeCloseTo(80 + 165, 6);
	});

	it('shows each half only the band of the words a piece owns', () => {
		const pieces = resolveMorphTextPieces(from, to, measure)!.pieces;
		const consectetur = pieces[1]!;
		expect(consectetur.fromBand.left).toBe(265);
		expect(consectetur.fromBand.bottom).toBe(24);
		expect(consectetur.fromBand.right).toBeGreaterThan(1000);
		expect(consectetur.fromBand.top).toBeLessThan(-1000);
		expect(consectetur.toBand.left).toBeLessThan(-1000);
		expect(consectetur.toBand.right).toBeCloseTo(126.5, 6);
		expect(consectetur.toBand.top).toBeCloseTo(26.4, 6);
		expect(consectetur.toBand.bottom).toBeCloseTo(52.8, 6);
	});

	it('lays each piece of one half exactly over the same words of the other', () => {
		for (const piece of resolveMorphTextPieces(from, to, measure)!.pieces) {
			const { from: a, to: b } = piece.alignment;
			const start = apply(morphTextIncomingStart(piece.alignment, to), to, b.x, b.y);
			expect(start.x).toBeCloseTo(a.x, 6);
			expect(start.y).toBeCloseTo(a.y, 6);
			const end = apply(morphTextGhostEnd(piece.alignment, from), from, a.x, a.y);
			expect(end.x).toBeCloseTo(b.x, 6);
			expect(end.y).toBeCloseTo(b.y, 6);
		}
	});

	it('grows a box to hold the lines its text overflows it by', () => {
		const short = textBox('short-1', 20, { height: 40 });
		const split = resolveMorphTextPieces(short, textBox('short-2', 22, { height: 40 }), measure)!;
		expect(split.fromHeight).toBe(72);
		expect(split.toHeight).toBe(80);
		expect(resolveMorphTextPieces(from, to, measure)!.toHeight).toBe(120);
	});

	it('refuses a layout it cannot reproduce with confidence', () => {
		// A break within a pixel of the column edge: 38 characters of 10.5 px.
		expect(
			resolveMorphTextPieces(textBox('edge-1', 21, { width: 399.5 }), to, measure),
		).toBeUndefined();
		// Different text, centred-anchored text and a word wider than its column.
		expect(resolveMorphTextPieces(from, { ...to, text: 'Other' } as PptxElement, measure)).toBe(
			undefined,
		);
		const middle = textBox('mid', 22);
		(middle as { textStyle: Record<string, unknown> }).textStyle.vAlign = 'middle';
		expect(resolveMorphTextPieces(from, middle, measure)).toBeUndefined();
		expect(
			resolveMorphTextPieces(from, textBox('narrow', 22, { width: 100 }), measure),
		).toBeUndefined();
	});

	it('paints the pieces as crossfade groups in place of the pair', () => {
		const plan = buildMorphTransitionPlan(slide([from]), slide([to]), 2000, 'object', {
			measureText: measure,
		})!;
		const ids = plan.crossfadeGroups.map((group) => group.incoming.id);
		expect(ids).toEqual([0, 1, 2, 3, 4].map((k) => morphTextPieceId('wrap-2', k)));
		expect(plan.crossfadeGroups[1]!.outgoing.id).toBe(morphTextPieceId('wrap-1', 1));
		expect(plan.textPieceSources?.get(morphTextPieceId('wrap-1', 1))).toBe('wrap-1');
		expect(plan.textPieceSources?.get(morphTextPieceId('wrap-2', 1))).toBe('wrap-2');
		// The pair's own ghost is not painted, and its stage copy is held invisible.
		expect(plan.outgoingElements.map((element) => element.id)).not.toContain('wrap-1');
		expect(plan.outgoingAnimations.has('wrap-1')).toBe(false);
		expect(plan.incomingAnimations.get('wrap-2')).toMatch(/^pptx-morph-lifted-hidden /u);
		// Each copy rides its own journey, clipped to its own band.
		const animation = plan.overlayIncomingAnimations.get(morphTextPieceId('wrap-2', 1))!;
		const name = animation.split(' ')[0]!;
		expect(plan.keyframesCss).toContain(`@keyframes ${name}`);
		expect(plan.keyframesCss).toMatch(/clip-path: polygon\(-2000px 26\.4px, 126\.5px 26\.4px/u);
		expect(plan.overlayStack.filter((entry) => entry.kind === 'crossfade')).toHaveLength(5);
	});

	it('leaves a pair the whole-box text mapping carries alone', () => {
		// The same layout moved elsewhere: one rigid move, no pieces.
		const moved = textBox('moved', 20, { x: 300, y: 300 });
		const plan = buildMorphTransitionPlan(slide([from]), slide([moved]), 2000, 'object', {
			measureText: measure,
		})!;
		expect(plan.textPieceSources).toBeUndefined();
		expect(plan.crossfadeGroups).toHaveLength(0);
	});
});

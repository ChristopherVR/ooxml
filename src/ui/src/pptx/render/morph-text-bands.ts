/**
 * The band of each half a morphing text piece shows (see `morph-text-pieces`).
 *
 * A piece is a copy of the whole box clipped to the words it owns. Cut along
 * the line boxes, the clip slices off whatever hangs past them (descenders,
 * italic overhangs, accents) and, when lines sit tight, shows the hanging
 * glyphs of the neighbouring line a second time. With the glyphs' ink measured
 * (`morph-text-ink`), the edge between two lines instead steps along the gap
 * between the ink of the line above and the ink of the line below: at the
 * line boxes' shared edge wherever that is clear, else as close to it as the
 * gap allows, and through the middle of the overlap where the two lines' glyph
 * boxes genuinely overlap. The pieces of neighbouring lines share that edge,
 * so they tile the box without a gap or a double.
 *
 * Words on one line part halfway between their ink (else their advances).
 *
 * @module render/morph-text-bands
 */
import type { MorphTextLayout, MorphTextRect, MorphTextWord } from './morph-text-lines';

/** One step of a band edge: at height `y` from `x0` to `x1`. */
export interface MorphTextEdgeStep {
	x0: number;
	x1: number;
	y: number;
}

/** A band: its bounding rectangle, and the stepped top and bottom edges when ink was measured. */
export interface MorphTextBand extends MorphTextRect {
	topEdge?: MorphTextEdgeStep[];
	bottomEdge?: MorphTextEdgeStep[];
}

/** How far an open band edge (a line's outer end, the first or last line) reaches. */
export const OPEN_EDGE_PX = 2000;
/** How far (px) an edge keeps clear of ink, for antialiasing. */
const INK_CLEARANCE_PX = 1;

/** Where two neighbouring words on one line part: halfway between their ink, else their advances. */
function between(left: MorphTextWord, right: MorphTextWord): number {
	return left.ink && right.ink ? (left.ink.right + right.ink.left) / 2 : (left.x1 + right.x0) / 2;
}

function glyphsOn(layout: MorphTextLayout, line: number): MorphTextRect[] | undefined {
	const rects: MorphTextRect[] = [];
	for (const word of layout.words) {
		if (word.line !== line) {
			continue;
		}
		if (!word.glyphs) {
			return undefined;
		}
		rects.push(...word.glyphs);
	}
	return rects;
}

/**
 * The edge between line `k` and line `k + 1`, as steps across the whole
 * width, or `undefined` without measured ink.
 */
function boundary(layout: MorphTextLayout, k: number): MorphTextEdgeStep[] | undefined {
	const above = glyphsOn(layout, k);
	const below = glyphsOn(layout, k + 1);
	const upper = layout.lines[k];
	const lower = layout.lines[k + 1];
	if (!above || !below || !upper || !lower) {
		return undefined;
	}
	const slot = (upper.top + upper.height + lower.top) / 2;
	const clear = INK_CLEARANCE_PX;
	const xs = [...above, ...below].flatMap((rect) => [rect.left - clear, rect.right + clear]);
	const cuts = [...new Set([-OPEN_EDGE_PX, ...xs, OPEN_EDGE_PX])].sort((a, b) => a - b);
	const steps: MorphTextEdgeStep[] = [];
	for (let i = 0; i + 1 < cuts.length; i++) {
		const x0 = cuts[i] as number;
		const x1 = cuts[i + 1] as number;
		const spans = (rect: MorphTextRect) => rect.right + clear > x0 && rect.left - clear < x1;
		const bottom = Math.max(-Infinity, ...above.filter(spans).map((rect) => rect.bottom));
		const top = Math.min(Infinity, ...below.filter(spans).map((rect) => rect.top));
		const y =
			bottom + clear <= top - clear
				? Math.min(Math.max(slot, bottom + clear), top - clear)
				: (bottom + top) / 2;
		const last = steps.at(-1);
		if (last && Math.abs(last.y - y) < 1e-6) {
			last.x1 = x1;
		} else {
			steps.push({ x0, x1, y });
		}
	}
	return steps;
}

/** Every edge between consecutive lines of `layout`. */
export function morphTextBoundaries(layout: MorphTextLayout): (MorphTextEdgeStep[] | undefined)[] {
	return layout.lines.slice(1).map((_, k) => boundary(layout, k));
}

function within(steps: MorphTextEdgeStep[], left: number, right: number): MorphTextEdgeStep[] {
	return steps
		.filter((step) => step.x1 > left && step.x0 < right)
		.map((step) => ({ x0: Math.max(step.x0, left), x1: Math.min(step.x1, right), y: step.y }));
}

/**
 * The band of `layout` a run of words `[first, last]` owns, open at the outer
 * edges.
 *
 * @param boundaries - The layout's {@link morphTextBoundaries}.
 */
export function morphTextBandOf(
	layout: MorphTextLayout,
	first: number,
	last: number,
	boundaries: readonly (MorphTextEdgeStep[] | undefined)[],
): MorphTextBand {
	const words = layout.words;
	const head = words[first] as MorphTextWord;
	const tail = words[last] as MorphTextWord;
	const before = words[first - 1];
	const after = words[last + 1];
	const line = layout.lines[head.line];
	const sameLine = (word: MorphTextWord | undefined): word is MorphTextWord =>
		word !== undefined && word.line === head.line;
	const left = sameLine(before) ? between(before, head) : -OPEN_EDGE_PX;
	const right = sameLine(after) ? between(tail, after) : OPEN_EDGE_PX;
	const band: MorphTextBand = {
		left,
		right,
		top: head.line === 0 ? -OPEN_EDGE_PX : (line?.top ?? 0),
		bottom:
			head.line === layout.lines.length - 1 ? OPEN_EDGE_PX : (line?.top ?? 0) + (line?.height ?? 0),
	};
	const above = head.line > 0 ? boundaries[head.line - 1] : undefined;
	const below = boundaries[head.line];
	if (above) {
		band.topEdge = within(above, left, right);
		band.top = Math.min(...band.topEdge.map((step) => step.y));
	}
	if (below) {
		band.bottomEdge = within(below, left, right);
		band.bottom = Math.max(...band.bottomEdge.map((step) => step.y));
	}
	return band;
}

/** The band's outline, clockwise from its top-left corner, as `[x, y]` points. */
export function morphTextBandOutline(band: MorphTextBand): [number, number][] {
	const top: [number, number][] = band.topEdge?.flatMap((step) => [
		[step.x0, step.y] as [number, number],
		[step.x1, step.y] as [number, number],
	]) ?? [
		[band.left, band.top],
		[band.right, band.top],
	];
	const bottom: [number, number][] = band.bottomEdge
		? [...band.bottomEdge]
				.reverse()
				.flatMap((step) => [
					[step.x1, step.y] as [number, number],
					[step.x0, step.y] as [number, number],
				])
		: [
				[band.right, band.bottom],
				[band.left, band.bottom],
			];
	return [...top, ...bottom];
}

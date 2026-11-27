/**
 * The Draw tab's pens, shared by every product: the kinds of pen, the default gallery, how a saved
 * gallery is validated and how each kind lays its stroke down.
 *
 * @module drawingml/ink-pens
 */

export type InkPenKind = 'pen' | 'pencil' | 'highlighter' | 'actionPen';

export interface InkPen {
	kind: InkPenKind;
	/** `#rrggbb`. */
	color: string;
	/** Stroke width in CSS pixels. */
	width: number;
}

export const INK_PEN_KINDS: readonly InkPenKind[] = ['pen', 'pencil', 'highlighter', 'actionPen'];
/** The most pens a gallery keeps. */
export const MAX_INK_PENS = 12;

const mm = (value: number): number => Math.round(((value * 96) / 25.4) * 100) / 100;

/** The gallery a new document starts with, as Office's: three pens, a pencil and a highlighter. */
export const DEFAULT_INK_PENS: readonly InkPen[] = [
	{ kind: 'pen', color: '#000000', width: mm(0.5) },
	{ kind: 'pen', color: '#E81123', width: mm(0.5) },
	{ kind: 'pen', color: '#0078D4', width: mm(0.5) },
	{ kind: 'pencil', color: '#404040', width: mm(1) },
	{ kind: 'highlighter', color: '#FFFF00', width: mm(6) },
];

/** Thickness choices (millimetres, as the Draw tab names them) and their CSS pixel widths. */
export const INK_PEN_THICKNESSES_MM: Readonly<Record<'pen' | 'highlighter', readonly number[]>> = {
	pen: [0.25, 0.35, 0.5, 0.75, 1, 2, 3.5],
	highlighter: [2, 4, 6, 8, 10],
};
export const inkThicknessPx = mm;
/** The widest pen the gallery holds: the widest highlighter (10 mm). */
export const MAX_INK_PEN_WIDTH = mm(10);

const COLOR = /^#[\da-f]{6}$/iu;

/** Whether `pen` is a pen the gallery can hold, with widths up to `maxWidth` pixels. */
export function isInkPen(pen: unknown, maxWidth = MAX_INK_PEN_WIDTH): pen is InkPen {
	if (!pen || typeof pen !== 'object') return false;
	const { kind, color, width } = pen as Partial<InkPen>;
	return (
		INK_PEN_KINDS.some((item) => item === kind) &&
		typeof color === 'string' &&
		COLOR.test(color) &&
		typeof width === 'number' &&
		width >= 0.5 &&
		width <= maxWidth
	);
}

/** A stored gallery reduced to its valid pens, at most {@link MAX_INK_PENS}. */
export function normalizeInkPens(value: unknown, maxWidth = MAX_INK_PEN_WIDTH): InkPen[] {
	if (!Array.isArray(value)) return [];
	return value
		.slice(0, MAX_INK_PENS)
		.flatMap((pen: unknown) =>
			isInkPen(pen, maxWidth) ? [{ kind: pen.kind, color: pen.color, width: pen.width }] : [],
		);
}

/** How opaque a pen's strokes are: a highlighter is translucent and a pencil a little grainy. */
export function inkPenOpacity(kind: InkPenKind): number {
	return kind === 'highlighter' ? 0.4 : kind === 'pencil' ? 0.72 : 1;
}

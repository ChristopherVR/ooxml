// The `a:gradFill` reader (ECMA-376 Part 1, 20.1.8.33), written once over a DrawingNodeView so
// the DOM fill reader and the pptx object-tree codec read a gradient the same way. The neutral
// fill keeps only what a resolved model needs; the read also returns every stop as written and the
// attributes some consumers keep beside the model (`flip`, `rotWithShape`).
import type { DrawingNodeView } from './node-view';
import type { DrawingColor, DrawingFill } from './types';

export type DrawingGradientFill = Extract<DrawingFill, { kind: 'gradient' }>;
export type DrawingRelativeRect = { l: number; t: number; r: number; b: number };

/** One `a:gs` as written: `position` is the raw `pos` (1/1000 percent), `color` its colour. */
export interface DrawingGradientStopEntry<N> {
	node: N;
	position?: number;
	color?: DrawingColor;
}

export interface DrawingGradientRead<N> {
	/** The neutral gradient: stops with a colour and a position, in document order. */
	fill: DrawingGradientFill;
	/** Every `a:gs` of `a:gsLst`, in document order, including ones the neutral fill drops. */
	stops: DrawingGradientStopEntry<N>[];
	/** Whether an `a:path` element is present (a path gradient even without `@path`). */
	hasPath: boolean;
	/** `@flip` as written (`ST_TileFlipMode`), trimmed and lower-cased. */
	flip?: string;
	/** `@rotWithShape` as a boolean when it is a valid `xsd:boolean`. */
	rotateWithShape?: boolean;
	/** `a:path/a:fillToRect` and `a:tileRect` with each absent side as 0 and unreadable sides kept as NaN. */
	fillToRectSides?: DrawingRelativeRect;
	tileRectSides?: DrawingRelativeRect;
}

/** An `xsd:boolean` attribute value. */
export function drawingBoolean(value: string | undefined): boolean | undefined {
	const token = value?.trim().toLowerCase();
	if (token === '1' || token === 'true') return true;
	if (token === '0' || token === 'false') return false;
	return undefined;
}

function relativeRect<N>(view: DrawingNodeView<N>, node: N | undefined) {
	if (node === undefined) return undefined;
	const side = (name: string) =>
		Number.parseInt(view.attribute(node, name)?.trim() || '0', 10) / 100000;
	return { l: side('l'), t: side('t'), r: side('r'), b: side('b') };
}

/** Reads `a:gradFill`. `sourceXml` is the caller's (a DOM reader serializes the element). */
export function readDrawingGradient<N>(
	view: DrawingNodeView<N>,
	gradFill: N,
): DrawingGradientRead<N> {
	const list = view.child(gradFill, 'gsLst') ?? gradFill;
	const stops = view.children(list, 'gs').map((node) => {
		const entry: DrawingGradientStopEntry<N> = { node };
		const position = Number.parseInt(view.attribute(node, 'pos') ?? '', 10);
		if (Number.isFinite(position)) entry.position = position;
		const color = view.color(node);
		if (color) entry.color = color;
		return entry;
	});
	const lin = view.child(gradFill, 'lin');
	const pathNode = view.child(gradFill, 'path');
	const angle = Number.parseInt((lin && view.attribute(lin, 'ang')) ?? '', 10);
	const scaled = lin ? view.attribute(lin, 'scaled') : undefined;
	const path = pathNode ? view.attribute(pathNode, 'path') : undefined;
	const fillToRectSides = relativeRect(
		view,
		pathNode === undefined ? undefined : view.child(pathNode, 'fillToRect'),
	);
	const tileRectSides = relativeRect(view, view.child(gradFill, 'tileRect'));
	const finite = (rect: DrawingRelativeRect | undefined) =>
		rect && Object.values(rect).every(Number.isFinite) ? rect : undefined;
	const fillToRect = finite(fillToRectSides);
	const tileRect = finite(tileRectSides);
	const fill: DrawingGradientFill = {
		kind: 'gradient',
		stops: stops.flatMap((stop) =>
			stop.color && stop.position !== undefined
				? [{ position: stop.position / 1000, color: stop.color }]
				: [],
		),
		...(Number.isFinite(angle) ? { angle: angle / 60000 } : {}),
		...(scaled === '1' || scaled === 'true' ? { scaled: true } : {}),
		...(scaled === '0' || scaled === 'false' ? { scaled: false } : {}),
		...(path ? { path } : {}),
		...(fillToRect ? { fillToRect: { ...fillToRect } } : {}),
		...(tileRect ? { tileRect: { ...tileRect } } : {}),
	};
	const read: DrawingGradientRead<N> = { fill, stops, hasPath: pathNode !== undefined };
	const flip = view.attribute(gradFill, 'flip')?.trim().toLowerCase();
	if (flip) read.flip = flip;
	const rotateWithShape = drawingBoolean(view.attribute(gradFill, 'rotWithShape'));
	if (rotateWithShape !== undefined) read.rotateWithShape = rotateWithShape;
	if (fillToRectSides) read.fillToRectSides = fillToRectSides;
	if (tileRectSides) read.tileRectSides = tileRectSides;
	return read;
}

/** Stops ordered by position; stops without a readable position sort last, as written. */
export function sortGradientStopEntries<N>(
	stops: readonly DrawingGradientStopEntry<N>[],
	missingPosition?: number,
): DrawingGradientStopEntry<N>[] {
	const key = (stop: DrawingGradientStopEntry<N>) => stop.position ?? missingPosition;
	return [...stops].sort((left, right) => {
		const a = key(left);
		const b = key(right);
		if (a === undefined && b === undefined) return 0;
		if (a === undefined) return 1;
		if (b === undefined) return -1;
		return a - b;
	});
}

/** The centre of a path gradient's focus rectangle, each axis clamped to 0..1. */
export function gradientFocusCenter(rect: DrawingRelativeRect): { x: number; y: number } {
	const clamp = (value: number) => Math.min(1, Math.max(0, value));
	return { x: clamp((rect.l + (1 - rect.r)) / 2), y: clamp((rect.t + (1 - rect.b)) / 2) };
}

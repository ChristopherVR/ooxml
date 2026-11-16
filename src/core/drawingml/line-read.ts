// The `a:ln` (CT_LineProperties, ECMA-376 Part 1, 20.1.2.2.24) field reader, written once over a
// DrawingNodeView so the DOM line reader (`parseDrawingLine`) and the pptx object-tree readers
// (text outlines, underline lines) read width, dash, cap, compound, alignment, join and ends the
// same way. The fill and the source XML are the caller's.
import type { DrawingNodeView } from './node-view';
import type { DrawingLine } from './types';

/** The line fields without fill and source XML. */
export type DrawingLineFields = Omit<DrawingLine, 'sourceXml' | 'fill'>;

/**
 * Reads the fields of `a:ln`. `sourceOf` serializes a custom dash stop for its `sourceXml`; without
 * it the stops carry none.
 */
export function readDrawingLineFields<N>(
	view: DrawingNodeView<N>,
	ln: N,
	sourceOf?: (node: N) => string,
): DrawingLineFields {
	const line: DrawingLineFields = {};
	const width = Number.parseInt(view.attribute(ln, 'w') ?? '', 10);
	if (Number.isFinite(width)) line.widthEmu = width;
	const prstDash = view.child(ln, 'prstDash');
	const dash = prstDash === undefined ? undefined : view.attribute(prstDash, 'val');
	if (dash) line.dash = dash;
	const cap = view.attribute(ln, 'cap');
	if (cap) line.cap = cap;
	const compound = view.attribute(ln, 'cmpd');
	if (compound !== undefined) line.compound = compound;
	const alignment = view.attribute(ln, 'algn');
	if (alignment !== undefined) line.alignment = alignment;
	const custom = view.child(ln, 'custDash');
	if (custom !== undefined && !dash) {
		line.customDash = view.children(custom, 'ds').flatMap((node) => {
			const dashLength = drawingDashLength(view.attribute(node, 'd') ?? null);
			const space = drawingDashLength(view.attribute(node, 'sp') ?? null);
			return dashLength !== undefined && space !== undefined
				? [{ dash: dashLength, space, ...(sourceOf ? { sourceXml: sourceOf(node) } : {}) }]
				: [];
		});
	}
	for (const kind of ['round', 'bevel', 'miter'] as const) {
		const join = view.child(ln, kind);
		if (join === undefined) continue;
		const limitText = view.attribute(join, 'lim');
		const limit = Number(limitText);
		line.join = { kind, ...(limitText !== undefined && Number.isFinite(limit) ? { limit } : {}) };
		break;
	}
	for (const end of ['headEnd', 'tailEnd'] as const) {
		const node = view.child(ln, end);
		if (node === undefined) continue;
		const value: NonNullable<DrawingLine['headEnd']> = {};
		for (const [attribute, key] of [
			['type', 'type'],
			['w', 'width'],
			['len', 'length'],
		] as const) {
			const text = view.attribute(node, attribute);
			if (text !== undefined) value[key] = text;
		}
		line[end] = value;
	}
	return line;
}

/** ST_PositivePercentage custom-dash bounds, matching the PowerPoint object-tree reader. */
export function drawingDashLength(value: string | null): number | undefined {
	if (!/^\d+$/u.test(value ?? '')) return undefined;
	const number = Number(value);
	return Number.isSafeInteger(number) && number <= 2147483647 ? number : undefined;
}

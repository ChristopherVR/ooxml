/**
 * The gradient stops the inspector's Fill section edits, read from and
 * written back to a `ShapeStyle`.
 *
 * Unlike the canvas's `sanitizeGradientStops`, these keep every field a stop
 * carries (its preserved `originalColorXml` and theme `colorRef`), so editing
 * one stop never drops the theme colours of the others. The list operations
 * are the format-neutral ones in `ooxml-core/chart` (sorted by position, at
 * least two stops).
 *
 * @module render/inspector/shape-fill-stops
 */
import { removeGradientStop, sortGradientStops, updateGradientStop } from 'ooxml-core/chart';
import type { ShapeStyle } from 'ooxml-core/pptx';
import { normalizeHexColor } from 'ooxml-core/color';

import { DEFAULT_FILL_COLOR } from '../../constants';

/** One stop of a shape's gradient fill, as the style stores it. */
export type ShapeFillStop = NonNullable<ShapeStyle['fillGradientStops']>[number];

const clampPosition = (position: number): number => Math.max(0, Math.min(100, position));

/**
 * The stops `style` paints, sorted by position: its own when it has at least
 * two usable ones, else two stops from its fill colour to white (what a shape
 * switching to a gradient starts with).
 */
export function shapeFillStopsOf(style: ShapeStyle | undefined): ShapeFillStop[] {
	const own = (style?.fillGradientStops ?? [])
		.filter(
			(stop) =>
				typeof stop?.color === 'string' &&
				stop.color.trim().length > 0 &&
				typeof stop.position === 'number' &&
				Number.isFinite(stop.position),
		)
		.map((stop) => ({ ...stop, position: clampPosition(stop.position) }));
	if (own.length >= 2) {
		return sortGradientStops(own);
	}
	return [
		{
			color: normalizeHexColor(style?.fillColor, DEFAULT_FILL_COLOR),
			position: 0,
			opacity: 1,
		},
		{ color: '#ffffff', position: 100, opacity: 1 },
	];
}

/** The style change that sets the gradient's stops to `stops`. */
export function shapeFillStopsPatch(stops: ShapeFillStop[]): Partial<ShapeStyle> {
	return { fillMode: 'gradient', fillGradientStops: stops };
}

/** Change stop `index` (re-sorted, so a moved stop keeps the list in order). */
export function editShapeFillStop(
	style: ShapeStyle | undefined,
	index: number,
	changes: Partial<ShapeFillStop>,
): Partial<ShapeStyle> | undefined {
	const stops = shapeFillStopsOf(style);
	if (index < 0 || index >= stops.length) {
		return undefined;
	}
	const next = { ...changes };
	if (next.position !== undefined) {
		next.position = clampPosition(next.position);
	}
	return shapeFillStopsPatch(updateGradientStop(stops, index, next));
}

/** Remove stop `index`; `undefined` when only two stops are left. */
export function removeShapeFillStop(
	style: ShapeStyle | undefined,
	index: number,
): Partial<ShapeStyle> | undefined {
	const next = removeGradientStop(shapeFillStopsOf(style), index);
	return next ? shapeFillStopsPatch(next) : undefined;
}

/**
 * Add a stop halfway between the last two stops, in the colour of the stop it
 * follows, fully opaque.
 */
export function addShapeFillStop(style: ShapeStyle | undefined): Partial<ShapeStyle> {
	const stops = shapeFillStopsOf(style);
	const before = stops[stops.length - 2];
	const after = stops[stops.length - 1];
	const position = Math.round(((before?.position ?? 0) + (after?.position ?? 100)) / 2);
	const stop: ShapeFillStop = {
		color: normalizeHexColor(before?.color, DEFAULT_FILL_COLOR),
		position,
		opacity: 1,
	};
	if (before?.colorRef) {
		stop.colorRef = before.colorRef;
	}
	return shapeFillStopsPatch(sortGradientStops([...stops, stop]));
}

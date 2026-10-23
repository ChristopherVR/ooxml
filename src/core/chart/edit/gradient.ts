// Gradient fill edits of a chart element or series (lifted from the xlsx edit layer, see
// PROVENANCE.md): create, preset, angle, geometry and per-stop position, colour, brightness and
// transparency, with the semantics of Excel's Format pane (two stops at least, colour edits keep the
// stop's transparency, an unchanged result is no edit). Stop colours arrive in the caller's own
// colour type `C` and are converted by `toColor` only when used, so validation order is the pane's.
import { withDrawingColorBrightness } from '../../drawingml/drawing-color-brightness';
import {
	withDrawingGradientGeometry,
	type DrawingGradientGeometryType,
	type RectGradientDirection,
} from '../../drawingml/gradient-geometry';
import { officeGradientPresetFill } from '../../drawingml/gradient-presets';
import type { DrawingColor, DrawingFill } from '../../drawingml/types';
import { removeGradientStop } from '../gradient-stop-edit';
import {
	assertChartPercent,
	drawingColorTransparency,
	withAlphaOf,
	withDrawingColorTransparency,
} from './paint';

/** One gradient edit; `C` is the colour type a stop colour is given in. */
export type ChartGradientFillEdit<C = DrawingColor> =
	| { kind: 'create' }
	| { kind: 'preset'; id: number }
	| { kind: 'angle'; value: number }
	| { kind: 'geometry'; type: DrawingGradientGeometryType; direction?: RectGradientDirection }
	| {
			kind: 'stop';
			index: number;
			position?: number;
			color?: C;
			transparency?: number;
			brightness?: number;
	  }
	| { kind: 'add'; index: number }
	| { kind: 'remove'; index: number };

export type ChartGradientFillValue = Extract<DrawingFill, { kind: 'gradient' }>;

/** Transparency of a gradient stop in percent (0 when the stop is missing). */
export function chartGradientStopTransparency(fill: ChartGradientFillValue, index: number): number {
	const color = fill.stops[index]?.color;
	return (color && drawingColorTransparency(color)) ?? 0;
}

const gradientPercent = (value: number) =>
	assertChartPercent(value, 'Gradient value must be from 0 to 100');

const identity = <C>(color: C) => color as unknown as DrawingColor;

/**
 * The gradient after `edit`, and the stop the editor should select, or `undefined` when the edit
 * does not apply (a stop edit on a non-gradient fill) or changes nothing. `defaultColor` is the
 * first stop of a gradient created from a non-gradient fill.
 */
export function editChartGradientFill<C = DrawingColor>(
	old: DrawingFill | undefined,
	edit: ChartGradientFillEdit<C>,
	defaultColor: DrawingColor,
	toColor: (color: C) => DrawingColor = identity,
): { fill: ChartGradientFillValue; stopIndex: number } | undefined {
	let fill: ChartGradientFillValue;
	if (edit.kind === 'preset') fill = officeGradientPresetFill(edit.id);
	else if (old?.kind === 'gradient') fill = structuredClone(old);
	else {
		if (edit.kind !== 'create') return undefined;
		fill = {
			kind: 'gradient',
			angle: 90,
			scaled: true,
			stops: [
				{ position: 0, color: structuredClone(defaultColor) },
				{ position: 100, color: { kind: 'scheme', value: 'lt1', transforms: [] } },
			],
		};
	}
	let stopIndex = 0;
	if (edit.kind === 'create' && old?.kind === 'gradient') return undefined;
	if (edit.kind === 'angle') {
		if (!Number.isFinite(edit.value) || edit.value < 0 || edit.value > 360)
			throw new RangeError('Gradient angle must be from 0 to 360');
		if (fill.path) return undefined;
		fill.angle = edit.value % 360;
	} else if (edit.kind === 'geometry') {
		if (edit.type === 'linear' && !fill.path) return undefined;
		if (edit.type !== 'linear' && fill.path === edit.type && edit.direction === undefined)
			return undefined;
		fill = withDrawingGradientGeometry(fill, edit.type, edit.direction);
	} else if (edit.kind !== 'create' && edit.kind !== 'preset') {
		const stop = fill.stops[edit.index];
		if (!Number.isInteger(edit.index) || !stop)
			throw new RangeError(`No gradient stop at index ${edit.index}`);
		if (edit.kind === 'remove') {
			const stops = removeGradientStop(fill.stops, edit.index);
			if (!stops) return undefined;
			fill.stops = stops;
			stopIndex = Math.min(edit.index, stops.length - 1);
		} else {
			let target = stop;
			if (edit.kind === 'add') {
				target = structuredClone(stop);
				target.position = Math.round(
					(stop.position + (fill.stops[edit.index + 1]?.position ?? 100)) / 2,
				);
				fill.stops.push(target);
			} else {
				if (edit.position !== undefined) {
					gradientPercent(edit.position);
					target.position = edit.position;
				}
				if (edit.color !== undefined) target.color = withAlphaOf(toColor(edit.color), target.color);
				if (edit.brightness !== undefined)
					target.color = withDrawingColorBrightness(target.color, edit.brightness);
				if (edit.transparency !== undefined) {
					gradientPercent(edit.transparency);
					target.color = withDrawingColorTransparency(target.color, edit.transparency);
				}
			}
			stopIndex = fill.stops.indexOf(target);
		}
	}
	if (JSON.stringify(fill) === JSON.stringify(old)) return undefined;
	return { fill, stopIndex };
}

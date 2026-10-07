import { removeGradientStop } from '../../chart/gradient-stop-edit';
import { chartPaletteSeriesColorChoice, findChartColorPalette } from '../../chart/color-palettes';
import { resolveDrawingColor } from '../../diagram/drawing-color';
import { withDrawingColorBrightness } from '../../diagram/drawing-color-brightness';
import { officeGradientPresetFill } from '../../diagram/gradient-presets';
import {
	withDrawingGradientGeometry,
	type RectGradientDirection,
} from '../../diagram/gradient-geometry';
import type { DiagramFill } from '../../diagram/types';
import type { ChartObject, Color } from '../model';
import type { ChartPatch } from './charts';
import { chartDrawingColor } from './chart-series-fill';
import { chartSeriesSolidColor } from './chart-series-transparency';

export type ChartGradientEdit =
	| { kind: 'create' }
	| { kind: 'preset'; id: number }
	| { kind: 'angle'; value: number }
	| { kind: 'geometry'; type: 'linear' | 'rect'; direction?: RectGradientDirection }
	| {
			kind: 'stop';
			index: number;
			position?: number;
			color?: Color;
			transparency?: number;
			brightness?: number;
	  }
	| { kind: 'add'; index: number }
	| { kind: 'remove'; index: number };

type Gradient = Extract<DiagramFill, { kind: 'gradient' }>;
export function chartGradientStopTransparency(fill: Gradient, index: number): number {
	const color = fill.stops[index]?.color;
	const value =
		color &&
		resolveDrawingColor(color, { scheme: () => '#000000' }, { transformOrder: 'document' });
	return value ? Math.round((1 - value.alpha) * 100000) / 1000 : 0;
}

function percent(value: number): void {
	if (!Number.isFinite(value) || value < 0 || value > 100)
		throw new RangeError('Gradient value must be from 0 to 100');
}

/** Native gradient edits retain geometry, source flags and unrelated chart properties. */
export function chartSeriesGradientPatch(
	chart: ChartObject,
	index: number,
	edit: ChartGradientEdit,
): { patch: ChartPatch; stopIndex: number } | undefined {
	const current = chart.series[index];
	if (!Number.isInteger(index) || !current)
		throw new RangeError(`No chart series at index ${index}`);
	const old = current.fill;
	let fill: Gradient;
	if (edit.kind === 'preset') fill = officeGradientPresetFill(edit.id);
	else if (old?.kind === 'gradient') fill = structuredClone(old);
	else {
		if (edit.kind !== 'create') return undefined;
		const color = structuredClone(
			chartSeriesSolidColor(chart, index) ??
				chartPaletteSeriesColorChoice(
					findChartColorPalette(chart.colorPalette ?? 10) ?? findChartColorPalette(10)!,
					index,
					chart.series.length,
				),
		);
		fill = {
			kind: 'gradient',
			angle: 90,
			scaled: true,
			stops: [
				{ position: 0, color },
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
		if (edit.type === 'rect' && fill.path === 'rect' && edit.direction === undefined)
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
					percent(edit.position);
					target.position = edit.position;
				}
				if (edit.color) {
					const color = chartDrawingColor(edit.color);
					if (!color) throw new RangeError('Invalid gradient color');
					color.transforms.push(
						...target.color.transforms.filter((item) =>
							['alpha', 'alphaMod', 'alphaOff'].includes(item.name),
						),
					);
					target.color = color;
				}
				if (edit.brightness !== undefined)
					target.color = withDrawingColorBrightness(target.color, edit.brightness);
				if (edit.transparency !== undefined) {
					percent(edit.transparency);
					target.color.transforms = target.color.transforms.filter(
						(item) => !['alpha', 'alphaMod', 'alphaOff'].includes(item.name),
					);
					if (edit.transparency !== 0)
						target.color.transforms.push({
							name: 'alpha',
							value: String(Math.round((100 - edit.transparency) * 1000)),
						});
				}
			}
			stopIndex = fill.stops.indexOf(target);
		}
	}
	if (JSON.stringify(fill) === JSON.stringify(old)) return undefined;
	const series = structuredClone(chart.series);
	const next = series[index]!;
	next.fill = fill;
	delete next.color;
	delete next.drawingColor;
	delete next.pointColors;
	delete next.pointFills;
	return { patch: { series }, stopIndex };
}

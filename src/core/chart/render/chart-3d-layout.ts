// The 3-D box of a cartesian chart: its size, the bands along the category axis, the rows the
// series occupy in depth and the value scale. Scene units are the plot rectangle's pixels before
// the camera fits the box into it.
import { clusteredBarGeometry, type ClusteredBarGeometry } from '../bar-cluster-geometry';
import type { Chart3DSize } from './chart-3d-camera';
import type { Chart3DViewInput } from './chart-3d-projection';
import { chartBarSpacing } from './chart-spacing';
import type { Rect } from './chart-svg-util';
import type { ChartViewModel } from './chart-view';

export interface Chart3DLayout {
	size: Chart3DSize;
	horizontal: boolean;
	/** Length of the category axis and of the value axis in scene units. */
	catLength: number;
	valueLength: number;
	band: number;
	/** Depth rows: one per series when the series stand behind each other, else one. */
	rows: number;
	cell: number;
	/** The depth a mark occupies inside its row. */
	markDepth: number;
	bars: ClusteredBarGeometry | undefined;
	/** Value to a position along the value axis, 0..valueLength. */
	value: (v: number) => number;
	/** Row `r` of the depth axis as [front, back]. */
	rowZ: (r: number) => [number, number];
}

const DEFAULT_GAP_DEPTH = 150;
const clamp = (v: number, lo: number, hi: number): number => Math.min(Math.max(v, lo), hi);

/** Whether the series of this chart stand in separate depth rows. */
export function chart3DSeriesInDepth(model: ChartViewModel): boolean {
	return model.type === 'surface' || model.grouping === 'standard';
}

export function chart3DLayout(
	model: ChartViewModel,
	inner: Rect,
	view: Chart3DViewInput,
): Chart3DLayout {
	const horizontal = model.horizontal;
	const w = Math.max(1, inner.w);
	const h = Math.max(
		1,
		inner.h * (view.hPercent !== undefined ? clamp(view.hPercent, 5, 500) / 100 : 1),
	);
	const catLength = horizontal ? h : w;
	const valueLength = horizontal ? w : h;
	const cats = Math.max(1, model.categories.length);
	const band = catLength / cats;
	const rows = chart3DSeriesInDepth(model) ? Math.max(1, model.series.length) : 1;
	const isBar = model.type === 'bar' || model.type === 'column';
	const clustered = isBar && model.grouping === 'clustered';
	const bars = isBar
		? clusteredBarGeometry(
				band,
				clustered ? Math.max(1, model.series.length) : 1,
				chartBarSpacing(model),
			)
		: undefined;
	const depthScale = clamp(view.depthPercent ?? 100, 20, 2000) / 100;
	const gap = clamp(view.gapDepth ?? DEFAULT_GAP_DEPTH, 0, 500) / 100;
	const wanted = (bars?.singleBarWidth ?? band * 0.5) * depthScale;
	const wantedCell = rows === 1 ? wanted * 1.5 : wanted * (1 + gap);
	const depth = Math.min(rows * wantedCell, Math.max(w, h) * 0.9);
	const cell = depth / rows;
	const scale = model.valueAxis;
	const span = scale ? scale.max - scale.min || 1 : 1;
	return {
		size: { w, h, d: depth },
		horizontal,
		catLength,
		valueLength,
		band,
		rows,
		cell,
		markDepth: rows === 1 ? cell / 1.5 : cell / (1 + gap),
		bars,
		value: (v) => ((v - (scale?.min ?? 0)) / span) * valueLength || 0,
		rowZ: (r) => {
			const mark = rows === 1 ? cell / 1.5 : cell / (1 + gap);
			const z0 = r * cell + (cell - mark) / 2;
			return [z0, z0 + mark];
		},
	};
}

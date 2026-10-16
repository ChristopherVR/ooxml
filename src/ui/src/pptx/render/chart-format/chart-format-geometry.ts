/**
 * Geometry the chart Format hit-test and renderer share: text boxes, the
 * plot rectangle and the axis an axis-label list belongs to.
 *
 * @module render/chart-format/chart-format-geometry
 */
import type { PptxChartData } from 'ooxml-core/pptx';

import type { ChartViewModel, SvgText } from '../chart-view-model-types';

/** Approximate box of an SVG text node (no layout available in pure code). */
export function chartTextBox(text: SvgText, pad = 3) {
	const width = Math.max(4, text.text.length * text.fontSize * 0.58);
	const left =
		text.textAnchor === 'start'
			? text.x
			: text.textAnchor === 'end'
				? text.x - width
				: text.x - width / 2;
	const middle = text.dominantBaseline === 'middle' || text.dominantBaseline === 'central';
	const top = middle ? text.y - text.fontSize / 2 : text.y - text.fontSize;
	return {
		left: left - pad,
		top: top - pad,
		right: left + width + pad,
		bottom: top + text.fontSize + pad,
	};
}

/** The plot rectangle, from the gridlines (or the axis label band when there are none). */
export function chartPlotBounds(vm: ChartViewModel) {
	const lines = [...vm.gridlines, ...(vm.secondaryGridlines ?? [])];
	const xs = lines.flatMap((l) => [l.x1, l.x2]);
	const ys = lines.flatMap((l) => [l.y1, l.y2]);
	// No gridlines: the data marks bound the plot (a chart drawn without them).
	for (const prim of vm.primitives) {
		if (prim.kind === 'rect' && prim.part) {
			xs.push(prim.x, prim.x + prim.w);
			ys.push(prim.y, prim.y + prim.h);
		} else if (prim.kind === 'circle' && prim.part) {
			xs.push(prim.cx - prim.r, prim.cx + prim.r);
			ys.push(prim.cy - prim.r, prim.cy + prim.r);
		}
	}
	if (xs.length === 0) {
		return undefined;
	}
	return {
		left: Math.min(...xs),
		right: Math.max(...xs),
		top: Math.min(...ys),
		bottom: Math.max(...ys),
	};
}

/** Index in `data.axes` of the n-th visible axis of `types`, or -1. */
export function chartAxisIndex(data: PptxChartData, types: readonly string[], nth = 0): number {
	let seen = 0;
	const axes = data.axes ?? [];
	for (let i = 0; i < axes.length; i++) {
		if (!axes[i]!.deleted && types.includes(axes[i]!.axisType)) {
			if (seen === nth) {
				return i;
			}
			seen++;
		}
	}
	return -1;
}

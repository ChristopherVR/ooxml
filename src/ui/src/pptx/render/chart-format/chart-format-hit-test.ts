/**
 * Which chart element a click on a rendered chart landed on, so a chart
 * Format pane (and Chart Format > Current Selection) can target it the way
 * PowerPoint does: a series, an axis, its gridlines, the legend, the title,
 * the data labels, the plot area or the chart area.
 *
 * Data marks and the title carry `data-chart-part` in every binding, so they
 * are read off the event path. Everything else is found geometrically from
 * the chart's view model (the same one each binding paints), which keeps
 * this decision in one place instead of tagging five renderers.
 *
 * @module render/chart-format/chart-format-hit-test
 */
import type { PptxChartData, PptxChartFormatTargetKey, PptxElement } from 'ooxml-core/pptx';

import { findChartPartElementInEvent } from '../chart-event-target';
import type { ChartPathEvent } from '../chart-event-target';
import { CHART_PART_ATTR, CHART_PART_SERIES_ATTR } from '../chart-interaction';
import { computeChartLegendLayout } from '../chart-legend-layout';
import { buildChartViewModel } from '../chart-view-model';
import type { ChartViewModel, SvgLine, SvgText } from '../chart-view-model';
import { chartAxisIndex, chartPlotBounds, chartTextBox } from './chart-format-geometry';

export { chartAxisIndex, chartPlotBounds, chartTextBox } from './chart-format-geometry';

/** A point in the chart's SVG (view-box) coordinates. */
export interface ChartPoint {
	x: number;
	y: number;
}

function inBox(p: ChartPoint, box: { left: number; top: number; right: number; bottom: number }) {
	return p.x >= box.left && p.x <= box.right && p.y >= box.top && p.y <= box.bottom;
}

/** A vertical axis' labels can be rotated; only plain, unrotated labels are boxed. */
function hitsText(p: ChartPoint, texts: readonly SvgText[] | undefined): boolean {
	return (texts ?? []).some((text) => !text.transform && inBox(p, chartTextBox(text)));
}

function nearLine(p: ChartPoint, line: SvgLine, tolerance: number): boolean {
	const dx = line.x2 - line.x1;
	const dy = line.y2 - line.y1;
	const length2 = dx * dx + dy * dy;
	const t =
		length2 === 0
			? 0
			: Math.max(0, Math.min(1, ((p.x - line.x1) * dx + (p.y - line.y1) * dy) / length2));
	const cx = line.x1 + t * dx;
	const cy = line.y1 + t * dy;
	return Math.hypot(p.x - cx, p.y - cy) <= tolerance;
}

/**
 * The element at `point` (SVG coordinates) of a chart whose view model is
 * `vm`. Never returns a target the model lacks (a missing axis falls through
 * to the plot area).
 */
export function chartFormatTargetAtPoint(
	vm: ChartViewModel,
	data: PptxChartData,
	point: ChartPoint,
): PptxChartFormatTargetKey {
	if (data.style?.hasLegend !== false && vm.legend.length > 0) {
		const hitLegend = computeChartLegendLayout(vm).some((item) =>
			inBox(point, {
				left: item.x - 3,
				top: item.y - 9,
				right: item.x + 16 + item.label.length * item.fontSize * 0.58,
				bottom: item.y + 6,
			}),
		);
		if (hitLegend) {
			return 'legend';
		}
	}
	if (data.style?.hasDataLabels && hitsText(point, vm.dataLabels)) {
		return 'dataLabels';
	}
	const valueAxis = chartAxisIndex(data, ['valAx']);
	const categoryAxis = chartAxisIndex(data, ['catAx', 'dateAx']);
	if (valueAxis >= 0 && hitsText(point, vm.axisLabels)) {
		return `axis:${valueAxis}`;
	}
	const secondaryValue = chartAxisIndex(data, ['valAx'], 1);
	if (secondaryValue >= 0 && hitsText(point, vm.secondaryAxisLabels)) {
		return `axis:${secondaryValue}`;
	}
	if (categoryAxis >= 0 && hitsText(point, vm.categoryLabels)) {
		return `axis:${categoryAxis}`;
	}
	if (valueAxis >= 0 && data.axes?.[valueAxis]?.majorGridlines !== false) {
		if (vm.gridlines.some((line) => nearLine(point, line, 3))) {
			return `majorGridlines:${valueAxis}`;
		}
	}
	const bounds = chartPlotBounds(vm);
	if (bounds && inBox(point, bounds)) {
		return 'plotArea';
	}
	return 'chartArea';
}

/** What a chart Format hit-test needs from a DOM event. */
export interface ChartFormatPointerEvent extends ChartPathEvent {
	readonly clientX: number;
	readonly clientY: number;
}

/** Map a client point into `svg`'s view-box coordinates. */
export function clientToChartPoint(
	svg: SVGSVGElement,
	clientX: number,
	clientY: number,
	vm: Pick<ChartViewModel, 'svgWidth' | 'svgHeight'>,
): ChartPoint {
	const ctm = typeof svg.getScreenCTM === 'function' ? svg.getScreenCTM() : null;
	if (ctm && typeof svg.createSVGPoint === 'function') {
		const p = svg.createSVGPoint();
		p.x = clientX;
		p.y = clientY;
		const local = p.matrixTransform(ctm.inverse());
		return { x: local.x, y: local.y };
	}
	const rect = svg.getBoundingClientRect();
	const scaleX = rect.width > 0 ? vm.svgWidth / rect.width : 1;
	const scaleY = rect.height > 0 ? vm.svgHeight / rect.height : 1;
	return { x: (clientX - rect.left) * scaleX, y: (clientY - rect.top) * scaleY };
}

/**
 * The tagged chart part under the pointer. A mark's pointerdown can capture the pointer, which
 * retargets the following click / double-click to the chart host, so the event path no longer
 * names the mark; the hit list at the pointer still does.
 */
function chartPartUnderPointer(
	event: ChartFormatPointerEvent,
	host: Element | null | undefined,
): Element | null {
	const doc = host?.ownerDocument;
	if (!host || !doc || typeof doc.elementsFromPoint !== 'function') {
		return null;
	}
	for (const node of doc.elementsFromPoint(event.clientX, event.clientY)) {
		if (!host.contains(node)) {
			continue;
		}
		const part = node.closest(`[${CHART_PART_ATTR}]`);
		if (part) {
			return part;
		}
	}
	return null;
}

/** The chart's `<svg>`, also inside a 3D chart view's open shadow root (its 2D chrome overlay). */
function findChartSvg(host: Element | null | undefined): SVGSVGElement | null {
	const direct = host?.querySelector('svg');
	if (direct) {
		return direct;
	}
	for (const node of host?.querySelectorAll('*') ?? []) {
		const nested = node.shadowRoot?.querySelector('svg');
		if (nested) {
			return nested;
		}
	}
	return null;
}

/**
 * The chart element an event on a rendered chart targets. `host` is the
 * chart's wrapper (any ancestor of its `<svg>`), `element` the chart element.
 * Every binding's click, double-click and context-menu handlers call this.
 */
export function chartFormatTargetFromEvent(
	event: ChartFormatPointerEvent,
	host: Element | null | undefined,
	element: PptxElement | null | undefined,
): PptxChartFormatTargetKey {
	if (!element || element.type !== 'chart' || !element.chartData) {
		return 'chartArea';
	}
	const part = findChartPartElementInEvent(event) ?? chartPartUnderPointer(event, host);
	const role = part?.getAttribute(CHART_PART_ATTR);
	if (role === 'title') {
		return 'title';
	}
	if (role === 'dataPoint' || role === 'series') {
		const series = Number(part?.getAttribute(CHART_PART_SERIES_ATTR));
		if (Number.isInteger(series) && element.chartData.series[series]) {
			return `series:${series}`;
		}
	}
	const svg = findChartSvg(host);
	if (!svg) {
		return 'chartArea';
	}
	const vm = buildChartViewModel(element);
	return chartFormatTargetAtPoint(
		vm,
		element.chartData,
		clientToChartPoint(svg, event.clientX, event.clientY, vm),
	);
}

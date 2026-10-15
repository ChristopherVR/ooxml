/**
 * Headless operations behind PowerPoint's chart Format task panes (Fill,
 * Border/Line and Effects of one chart element). Each is pure: it returns a
 * new `PptxChartData` and leaves its input alone, so a UI can apply it as one
 * undoable element patch.
 *
 * An edit is recorded in `PptxChartData.elementFormats` (what the save path
 * writes into the element's `c:spPr`) and mirrored onto the flat model
 * fields the renderer already reads (`style.chartAreaFill`, a series'
 * `color` / `lineWidth`, an axis' `majorGridlinesSpPr`, ...), so the change
 * shows on the slide at once.
 *
 * @module sdk/chart-element-format-operations
 */
import type {
	PptxChartData,
	PptxChartElementEffects,
	PptxChartElementFill,
	PptxChartElementFormat,
	PptxChartElementLine,
	PptxChartFormatTarget,
	PptxChartGradientFill,
	PptxChartSeries,
	PptxChartShapeProps,
} from '../../types';
import { isLineDrawnChartType } from '../../utils/chart-container-type-map';
import { chartFormatTargetKey } from '../../utils/chart-element-format-keys';
import { readChartElementFormat } from './chart-element-format-read';
import type { ChartFormatFillKind } from './chart-element-format-read';

export { readChartElementFormat, chartFormatTargetSupports } from './chart-element-format-read';
export type {
	ChartElementFormatView,
	ChartFormatFillKind,
	ChartFormatTargetSupport,
} from './chart-element-format-read';

/** What a pane edit changes; omitted sections and fields stay as they are. */
export interface ChartElementFormatPatch {
	fill?: Partial<Omit<PptxChartElementFill, 'kind'>> & {
		kind?: ChartFormatFillKind;
		gradient?: PptxChartGradientFill;
	};
	line?: Partial<PptxChartElementLine>;
	effects?: PptxChartElementEffects;
}

function withSeries(
	data: PptxChartData,
	index: number,
	update: (series: PptxChartSeries) => PptxChartSeries,
): PptxChartData {
	if (!data.series[index]) {
		return data;
	}
	return { ...data, series: data.series.map((s, i) => (i === index ? update(s) : s)) };
}

function shapeFromFormat(
	base: PptxChartShapeProps | undefined,
	fill: PptxChartElementFill | undefined,
	line: PptxChartElementLine | undefined,
): PptxChartShapeProps | undefined {
	const next: PptxChartShapeProps = { ...base };
	if (fill) {
		delete next.fillColor;
		if (fill.kind === 'solid' && fill.color) {
			next.fillColor = fill.color;
		}
	}
	if (line) {
		delete next.strokeColor;
		delete next.lineNoFill;
		if (line.kind === 'solid' && line.color) {
			next.strokeColor = line.color;
		} else if (line.kind === 'none') {
			next.lineNoFill = true;
		}
		if (line.width !== undefined) {
			next.strokeWidth = line.width;
		}
		if (line.dash !== undefined) {
			next.strokeDashStyle = line.dash;
		}
	}
	return Object.keys(next).length > 0 ? next : undefined;
}

/** Mirror a recorded format onto the flat fields the renderer reads. */
function mirrorFlatFields(
	data: PptxChartData,
	target: PptxChartFormatTarget,
	fill: PptxChartElementFill | undefined,
	line: PptxChartElementLine | undefined,
): PptxChartData {
	const paint = (p: PptxChartElementFill | PptxChartElementLine | undefined) =>
		!p || p.kind === 'automatic' ? undefined : p.kind === 'none' ? 'none' : p.color;
	switch (target.kind) {
		case 'chartArea':
		case 'plotArea': {
			const style = { ...data.style };
			const prefix = target.kind === 'chartArea' ? 'chartArea' : 'plotArea';
			if (fill) {
				delete style[`${prefix}Gradient`];
				const value = paint(fill);
				if (value === undefined) {
					delete style[`${prefix}Fill`];
				} else {
					style[`${prefix}Fill`] = value;
				}
			}
			if (line) {
				const value = paint(line);
				if (value === undefined) {
					delete style[`${prefix}Border`];
				} else {
					style[`${prefix}Border`] = value;
				}
			}
			return { ...data, style };
		}
		case 'title':
			return {
				...data,
				style: {
					...data.style,
					titleSpPr: shapeFromFormat(data.style?.titleSpPr ?? undefined, fill, line) ?? null,
				},
			};
		case 'axis':
		case 'majorGridlines':
		case 'minorGridlines': {
			const field =
				target.kind === 'axis'
					? 'spPr'
					: target.kind === 'majorGridlines'
						? 'majorGridlinesSpPr'
						: 'minorGridlinesSpPr';
			const axes = (data.axes ?? []).map((axis, i) =>
				i === target.axisIndex
					? { ...axis, [field]: shapeFromFormat(axis[field], undefined, line) }
					: axis,
			);
			return { ...data, axes };
		}
		case 'series':
			return withSeries(data, target.seriesIndex, (series) => {
				const next = { ...series };
				const lineDrawn = isLineDrawnChartType(series.seriesChartType ?? data.chartType);
				if (fill && !lineDrawn) {
					delete next.gradientFill;
					delete next.color;
					if (fill.kind === 'solid' && fill.color) {
						next.color = fill.color;
					}
				}
				if (line) {
					if (lineDrawn) {
						delete next.lineNoFill;
						if (line.kind === 'solid' && line.color) {
							next.color = line.color;
						} else if (line.kind === 'none') {
							next.lineNoFill = true;
						}
					}
					if (line.width !== undefined) {
						next.lineWidth = line.width;
					}
					if (line.dash !== undefined) {
						next.lineDashStyle = line.dash;
					}
				}
				return next;
			});
		default:
			return data;
	}
}

/** Apply a gradient fill through the existing gradient fields (chart area, plot area, series). */
function applyGradient(
	data: PptxChartData,
	target: PptxChartFormatTarget,
	gradient: PptxChartGradientFill,
): PptxChartData {
	if (target.kind === 'chartArea' || target.kind === 'plotArea') {
		const style = { ...data.style };
		style[`${target.kind}Gradient`] = gradient;
		delete style[`${target.kind}Fill`];
		return { ...data, style };
	}
	if (target.kind === 'series') {
		return withSeries(data, target.seriesIndex, (series) => ({
			...series,
			gradientFill: gradient,
		}));
	}
	return data;
}

/**
 * Apply a Format-pane edit to one chart element.
 *
 * @example
 * ```ts
 * const next = withChartElementFormat(chart.chartData, { kind: 'legend' }, {
 *   fill: { kind: 'solid', color: '#FFF2CC', transparency: 20 },
 *   line: { kind: 'solid', color: '#BF9000', width: 1.5, dash: 'dash' },
 * });
 * ```
 */
export function withChartElementFormat(
	data: PptxChartData,
	target: PptxChartFormatTarget,
	patch: ChartElementFormatPatch,
): PptxChartData {
	const key = chartFormatTargetKey(target);
	const view = readChartElementFormat(data, target);
	const previous: PptxChartElementFormat = data.elementFormats?.[key] ?? {};
	const entry: PptxChartElementFormat = { ...previous };
	let next = data;
	let fill: PptxChartElementFill | undefined;
	if (patch.fill && view.support.fill) {
		const kind = patch.fill.kind ?? view.fill.kind;
		if (kind === 'gradient') {
			if (view.support.gradient) {
				const gradient = patch.fill.gradient ?? view.fill.gradient;
				if (gradient) {
					next = applyGradient(next, target, gradient);
					delete entry.fill;
				}
			}
		} else {
			fill = {
				kind,
				...((patch.fill.color ?? view.fill.color)
					? { color: patch.fill.color ?? view.fill.color }
					: {}),
				...((patch.fill.transparency ?? view.fill.transparency) !== undefined
					? { transparency: patch.fill.transparency ?? view.fill.transparency }
					: {}),
			};
			entry.fill = fill;
		}
	}
	let line: PptxChartElementLine | undefined;
	if (patch.line) {
		line = { ...view.line, ...patch.line };
		entry.line = line;
	}
	if (patch.effects) {
		entry.effects = { ...view.effects, ...patch.effects };
	}
	next = mirrorFlatFields(next, target, fill, line);
	return { ...next, elementFormats: { ...next.elementFormats, [key]: entry } };
}

/**
 * PowerPoint's Reset to Match Style: drop the element's own formatting so the
 * chart style draws it again.
 */
export function resetChartElementFormat(
	data: PptxChartData,
	target: PptxChartFormatTarget,
): PptxChartData {
	const key = chartFormatTargetKey(target);
	let next: PptxChartData = data;
	switch (target.kind) {
		case 'chartArea':
		case 'plotArea': {
			const style = { ...data.style };
			delete style[`${target.kind}Fill`];
			delete style[`${target.kind}Gradient`];
			delete style[`${target.kind}Border`];
			next = { ...data, style };
			break;
		}
		case 'title':
			next = { ...data, style: { ...data.style, titleSpPr: null } };
			break;
		case 'axis':
		case 'majorGridlines':
		case 'minorGridlines': {
			const field =
				target.kind === 'axis'
					? 'spPr'
					: target.kind === 'majorGridlines'
						? 'majorGridlinesSpPr'
						: 'minorGridlinesSpPr';
			next = {
				...data,
				axes: (data.axes ?? []).map((axis, i) => {
					if (i !== target.axisIndex) {
						return axis;
					}
					const copy = { ...axis };
					delete copy[field];
					return copy;
				}),
			};
			break;
		}
		case 'series':
			next = withSeries(data, target.seriesIndex, (series) => {
				const copy = { ...series };
				delete copy.color;
				delete copy.gradientFill;
				delete copy.lineWidth;
				delete copy.lineDashStyle;
				delete copy.lineNoFill;
				return copy;
			});
			break;
		default:
			break;
	}
	return { ...next, elementFormats: { ...next.elementFormats, [key]: { reset: true } } };
}

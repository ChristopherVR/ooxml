/**
 * Reading the current Fill / Line / Effects of one chart element for a chart
 * Format pane: the recorded pane edits over the flat model fields the chart
 * was loaded with.
 *
 * @module sdk/chart-element-format-read
 */
import type {
	PptxChartData,
	PptxChartElementEffects,
	PptxChartElementLine,
	PptxChartFormatTarget,
	PptxChartGradientFill,
	PptxChartShapeProps,
} from '../../types';
import { isLineDrawnChartType } from '../../utils/chart-container-type-map';
import { chartFormatTargetKey } from '../../utils/chart-element-format-keys';

/** Fill kinds a pane offers. `gradient` exists only where the model has a gradient field. */
export type ChartFormatFillKind = 'none' | 'solid' | 'gradient' | 'automatic';

/** Which sections apply to a target. */
export interface ChartFormatTargetSupport {
	/** False for lines (axis, gridlines, a line series), which have a Line section only. */
	fill: boolean;
	/** Gradient fill is available (chart area, plot area, a filled series). */
	gradient: boolean;
}

/** The element's current formatting as a pane shows it. */
export interface ChartElementFormatView {
	support: ChartFormatTargetSupport;
	fill: {
		kind: ChartFormatFillKind;
		color?: string;
		transparency?: number;
		gradient?: PptxChartGradientFill;
	};
	line: PptxChartElementLine;
	effects: PptxChartElementEffects;
	/** Reset to Match Style was used and nothing has been set since. */
	reset: boolean;
}

/** The sections `target` supports in `data`. */
export function chartFormatTargetSupports(
	data: PptxChartData,
	target: PptxChartFormatTarget,
): ChartFormatTargetSupport {
	switch (target.kind) {
		case 'axis':
		case 'majorGridlines':
		case 'minorGridlines':
			return { fill: false, gradient: false };
		case 'series': {
			const series = data.series[target.seriesIndex];
			const lineDrawn = isLineDrawnChartType(series?.seriesChartType ?? data.chartType);
			return { fill: !lineDrawn, gradient: !lineDrawn };
		}
		case 'chartArea':
		case 'plotArea':
			return { fill: true, gradient: true };
		default:
			return { fill: true, gradient: false };
	}
}

/** A light-to-colour gradient, PowerPoint's first preset when switching to Gradient fill. */
function defaultGradient(color: string | undefined): PptxChartGradientFill {
	const base = color && /^#?[0-9a-f]{6}$/iu.test(color) ? color.replace(/^#?/u, '#') : '#4472C4';
	return {
		type: 'linear',
		angle: 90,
		stops: [
			{ color: '#FFFFFF', position: 0 },
			{ color: base.toUpperCase(), position: 100 },
		],
	};
}

function paintKind(paint: string | undefined): { kind: ChartFormatFillKind; color?: string } {
	if (paint === undefined) {
		return { kind: 'automatic' };
	}
	return paint === 'none' ? { kind: 'none' } : { kind: 'solid', color: paint };
}

function lineFromShape(shape: PptxChartShapeProps | null | undefined): PptxChartElementLine {
	if (!shape) {
		return { kind: 'automatic' };
	}
	const line: PptxChartElementLine = shape.lineNoFill
		? { kind: 'none' }
		: shape.strokeColor
			? { kind: 'solid', color: shape.strokeColor }
			: { kind: 'automatic' };
	if (shape.strokeWidth !== undefined) {
		line.width = shape.strokeWidth;
	}
	if (shape.strokeDashStyle) {
		line.dash = shape.strokeDashStyle;
	}
	return line;
}

/** The flat-field view of `target`, before recorded pane edits. */
function flatView(
	data: PptxChartData,
	target: PptxChartFormatTarget,
): Pick<ChartElementFormatView, 'fill' | 'line'> {
	const style = data.style;
	switch (target.kind) {
		case 'chartArea':
		case 'plotArea': {
			const gradient = style?.[`${target.kind}Gradient` as const];
			const fill = gradient
				? { kind: 'gradient' as const, gradient }
				: paintKind(style?.[`${target.kind}Fill` as const]);
			const border = paintKind(style?.[`${target.kind}Border` as const]);
			return {
				fill,
				line: border.kind === 'gradient' ? { kind: 'automatic' } : (border as PptxChartElementLine),
			};
		}
		case 'title': {
			const shape = style?.titleSpPr;
			return {
				fill: shape?.fillColor ? { kind: 'solid', color: shape.fillColor } : { kind: 'automatic' },
				line: lineFromShape(shape),
			};
		}
		case 'axis':
		case 'majorGridlines':
		case 'minorGridlines': {
			const axis = data.axes?.[target.axisIndex];
			const shape =
				target.kind === 'axis'
					? axis?.spPr
					: target.kind === 'majorGridlines'
						? axis?.majorGridlinesSpPr
						: axis?.minorGridlinesSpPr;
			return { fill: { kind: 'automatic' }, line: lineFromShape(shape) };
		}
		case 'series': {
			const series = data.series[target.seriesIndex];
			if (!series) {
				return { fill: { kind: 'automatic' }, line: { kind: 'automatic' } };
			}
			const lineDrawn = isLineDrawnChartType(series.seriesChartType ?? data.chartType);
			const line: PptxChartElementLine = series.lineNoFill
				? { kind: 'none' }
				: lineDrawn && series.color
					? { kind: 'solid', color: series.color }
					: { kind: 'automatic' };
			if (series.lineWidth !== undefined) {
				line.width = series.lineWidth;
			}
			if (series.lineDashStyle) {
				line.dash = series.lineDashStyle;
			}
			const fill = series.gradientFill
				? { kind: 'gradient' as const, gradient: series.gradientFill }
				: !lineDrawn && series.color
					? { kind: 'solid' as const, color: series.color }
					: { kind: 'automatic' as const };
			return { fill, line };
		}
		default:
			return { fill: { kind: 'automatic' }, line: { kind: 'automatic' } };
	}
}

/** The current Fill / Line / Effects of `target` as a Format pane shows them. */
export function readChartElementFormat(
	data: PptxChartData,
	target: PptxChartFormatTarget,
): ChartElementFormatView {
	const support = chartFormatTargetSupports(data, target);
	const recorded = data.elementFormats?.[chartFormatTargetKey(target)];
	const flat = recorded?.reset
		? { fill: { kind: 'automatic' as const }, line: { kind: 'automatic' as const } }
		: flatView(data, target);
	// A recorded fill/line wins, except where a later flat edit (a Chart Styles or
	// Change Colors pick) set a gradient the record does not describe.
	const fill =
		recorded?.fill && flat.fill.kind !== 'gradient' ? { ...recorded.fill } : { ...flat.fill };
	const line: PptxChartElementLine = recorded?.line ? { ...recorded.line } : { ...flat.line };
	if (fill.kind === 'solid' && flat.fill.kind === 'solid' && flat.fill.color) {
		fill.color = flat.fill.color;
	}
	const gradient =
		'gradient' in fill && fill.gradient ? fill.gradient : defaultGradient(fill.color);
	return {
		support,
		fill: { ...fill, ...(support.gradient ? { gradient } : {}) },
		line,
		effects: { ...recorded?.effects },
		reset: Boolean(recorded?.reset && !recorded.fill && !recorded.line && !recorded.effects),
	};
}

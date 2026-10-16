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
	PptxChartElementText,
	PptxChartElementThreeD,
	PptxChartPatternFill,
	PptxChartPictureFill,
	PptxChartFormatTarget,
	PptxChartGradientFill,
	PptxChartShapeProps,
} from '../../types';
import { isLineDrawnChartType } from '../../utils/chart-container-type-map';
import { chartFormatTargetKey } from '../../utils/chart-element-format-keys';

/** Fill kinds a pane offers. */
export type ChartFormatFillKind =
	| 'none'
	| 'solid'
	| 'gradient'
	| 'pattern'
	| 'picture'
	| 'automatic';

/** Which sections apply to a target. */
export interface ChartFormatTargetSupport {
	/** False for lines (axis, gridlines, a line series), which have a Line section only. */
	fill: boolean;
	/** Gradient fill is available (every filled element). */
	gradient: boolean;
	/** Pattern and picture fills are available (every filled element). */
	pattern: boolean;
	picture: boolean;
	/** The Effects tab has a 3-D Format section. */
	threeD: boolean;
	/** The element has text, so the pane offers Text Options. */
	text: boolean;
}

/** The element's current formatting as a pane shows it. */
export interface ChartElementFormatView {
	support: ChartFormatTargetSupport;
	fill: {
		kind: ChartFormatFillKind;
		color?: string;
		transparency?: number;
		gradient?: PptxChartGradientFill;
		pattern?: PptxChartPatternFill;
		picture?: PptxChartPictureFill;
	};
	line: PptxChartElementLine;
	effects: PptxChartElementEffects;
	threeD: PptxChartElementThreeD;
	text: PptxChartElementText;
	/** Reset to Match Style was used and nothing has been set since. */
	reset: boolean;
}

/** The sections `target` supports in `data`. */
export function chartFormatTargetSupports(
	data: PptxChartData,
	target: PptxChartFormatTarget,
): ChartFormatTargetSupport {
	const filled = (fill: boolean, threeD: boolean, text: boolean): ChartFormatTargetSupport => ({
		fill,
		gradient: fill,
		pattern: fill,
		picture: fill,
		threeD,
		text,
	});
	switch (target.kind) {
		case 'axis':
			return filled(false, false, true);
		case 'majorGridlines':
		case 'minorGridlines':
			return filled(false, false, false);
		case 'series': {
			const series = data.series[target.seriesIndex];
			const lineDrawn = isLineDrawnChartType(series?.seriesChartType ?? data.chartType);
			return filled(!lineDrawn, true, false);
		}
		case 'plotArea':
			return filled(true, true, false);
		default:
			// Chart area (its text options cover every chart text), title, legend, data labels.
			return filled(true, true, true);
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

/** PowerPoint's first pattern (5%) in the element's colour on white. */
function defaultPattern(color: string | undefined): PptxChartPatternFill {
	const base = color && /^#?[0-9a-f]{6}$/iu.test(color) ? color.replace(/^#?/u, '#') : '#4472C4';
	return { preset: 'pct5', foreground: base.toUpperCase(), background: '#FFFFFF' };
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
	const pattern = 'pattern' in fill && fill.pattern ? fill.pattern : defaultPattern(fill.color);
	if (line.kind === 'gradient' && !line.gradient) {
		line.gradient = defaultGradient(line.color);
	}
	return {
		support,
		fill: {
			...fill,
			...(support.gradient ? { gradient } : {}),
			...(support.pattern ? { pattern } : {}),
		},
		line,
		effects: { ...recorded?.effects },
		threeD: { ...recorded?.threeD },
		text: { ...recorded?.text },
		reset: Boolean(
			recorded?.reset &&
			!recorded.fill &&
			!recorded.line &&
			!recorded.effects &&
			!recorded.threeD &&
			!recorded.text,
		),
	};
}

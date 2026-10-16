/**
 * How a chart Format pane edit shows on the slide at once: the recorded
 * fill and line mirrored onto the flat model fields the renderer already
 * reads (`style.chartAreaFill`, a series' `color` / `lineWidth` /
 * `impliedPicture`, an axis' `majorGridlinesSpPr`, ...), and gradients
 * routed through the model's own gradient fields where it has them.
 *
 * @module sdk/chart-element-format-mirror
 */
import type {
	PptxChartData,
	PptxChartElementFill,
	PptxChartElementLine,
	PptxChartElementText,
	PptxChartFormatTarget,
	PptxChartGradientFill,
	PptxChartSeries,
	PptxChartShapeProps,
} from '../../types';
import { isLineDrawnChartType } from '../../utils/chart-container-type-map';

export function withSeries(
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
		} else if (line.kind === 'gradient' && line.gradient?.stops.length) {
			// A plain line where the gradient is not drawn: its last stop.
			next.strokeColor = line.gradient.stops[line.gradient.stops.length - 1]!.color;
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
export function mirrorFlatFields(
	data: PptxChartData,
	target: PptxChartFormatTarget,
	fill: PptxChartElementFill | undefined,
	line: PptxChartElementLine | undefined,
): PptxChartData {
	const paint = (p: PptxChartElementFill | PptxChartElementLine | undefined) =>
		!p || (p.kind !== 'none' && p.kind !== 'solid')
			? undefined
			: p.kind === 'none'
				? 'none'
				: p.color;
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
					delete next.impliedPicture;
					if (fill.kind === 'solid' && fill.color) {
						next.color = fill.color;
					} else if (fill.kind === 'picture' && fill.picture?.src) {
						// Render-only, like a loaded bare `a:blipFill` (never serialized from here).
						next.impliedPicture = {
							imageUrl: fill.picture.src,
							pictureFormat: fill.picture.tile ? 'stack' : 'stretch',
							...(fill.picture.transparency
								? { opacity: 1 - fill.picture.transparency / 100 }
								: {}),
						};
					}
				}
				if (line) {
					if (lineDrawn) {
						delete next.lineNoFill;
						delete next.lineGradientFill;
						if (line.kind === 'gradient' && line.gradient) {
							next.lineGradientFill = line.gradient;
						}
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

/** Whether the model has a gradient field of its own for `target`. */
export function hasGradientField(target: PptxChartFormatTarget): boolean {
	return target.kind === 'chartArea' || target.kind === 'plotArea' || target.kind === 'series';
}

/** Apply a gradient fill through the existing gradient fields (chart area, plot area, series). */
export function applyGradient(
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
 * Mirror a Text Options fill onto the flat text colour the renderer already
 * reads for `target` (the chart-wide text style, the title colour and runs,
 * the legend text style, an axis' label colour, the data labels' text
 * style). A solid fill sets it; Automatic clears it; other kinds leave it.
 */
export function mirrorTextColor(
	data: PptxChartData,
	target: PptxChartFormatTarget,
	text: PptxChartElementText | undefined,
): PptxChartData {
	const fill = text?.fill;
	if (!fill || (fill.kind !== 'solid' && fill.kind !== 'automatic')) {
		return data;
	}
	const color = fill.kind === 'solid' ? fill.color : undefined;
	const withColor = <T extends { color?: string }>(style: T | undefined): T => {
		const next = { ...style } as T;
		if (color) next.color = color;
		else delete next.color;
		return next;
	};
	const style = { ...data.style };
	switch (target.kind) {
		case 'chartArea':
			style.textStyle = withColor(style.textStyle);
			return { ...data, style };
		case 'title': {
			if (color) style.titleFontColor = color;
			else delete style.titleFontColor;
			const titleRuns = data.titleRuns?.map((run) => withColor(run));
			return { ...data, style, ...(titleRuns ? { titleRuns } : {}) };
		}
		case 'legend':
			style.legendTextStyle = withColor(style.legendTextStyle);
			return { ...data, style };
		case 'axis':
			return {
				...data,
				axes: (data.axes ?? []).map((axis, i) => {
					if (i !== target.axisIndex) return axis;
					const next = { ...axis };
					if (color) next.fontColor = color;
					else delete next.fontColor;
					return next;
				}),
			};
		case 'dataLabels':
			if (target.seriesIndex === undefined) {
				style.dataLabels = { ...style.dataLabels, txPr: withColor(style.dataLabels?.txPr) };
				return { ...data, style };
			}
			return withSeries(data, target.seriesIndex, (series) => ({
				...series,
				dataLabelOptions: {
					...series.dataLabelOptions,
					txPr: withColor(series.dataLabelOptions?.txPr),
				},
			}));
		default:
			return data;
	}
}

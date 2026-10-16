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

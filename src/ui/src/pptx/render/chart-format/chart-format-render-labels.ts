/**
 * Data-label formats on the slide: the label box fill, transparency and
 * border (through the existing label decorations), and the labels' shadow
 * and glow, cast by each label box, or by the label text when the labels
 * have no box.
 *
 * @module render/chart-format/chart-format-render-labels
 */
import type {
	PptxChartData,
	PptxChartDataLabelOptions,
	PptxChartElementFill,
	PptxChartElementFormat,
	PptxChartShapeProps,
} from 'ooxml-core/pptx';

import type { ChartViewModel } from '../chart-view-model-types';
import { effectsBehind, opacityOf, safeColor } from './chart-format-effects-svg';
import { chartTextBox } from './chart-format-geometry';

const HEX = /^#[0-9a-f]{6}$/iu;

/** Fill kinds a box paints itself (gradients of the chart and plot area have model fields). */
export const paints = (fill: PptxChartElementFill | undefined, ownGradient: boolean) =>
	fill?.kind === 'solid' ||
	fill?.kind === 'pattern' ||
	fill?.kind === 'picture' ||
	(fill?.kind === 'gradient' && ownGradient);

/** `#RRGGBB` at `opacity`, as an `rgba()` the SVG `fill` attribute accepts. */
function withOpacity(color: string, opacity: number): string {
	if (opacity >= 1 || !HEX.test(color)) {
		return color;
	}
	const n = Number.parseInt(color.slice(1), 16);
	return `rgba(${(n >> 16) & 255}, ${(n >> 8) & 255}, ${n & 255}, ${Math.round(opacity * 1000) / 1000})`;
}

/** The label box a data-label format paints, over the parsed one. */
function labelShapeFrom(
	format: PptxChartElementFormat | undefined,
	current: PptxChartShapeProps | undefined,
): PptxChartShapeProps | undefined {
	const fill = format?.fill;
	const line = format?.line;
	if (fill?.kind !== 'solid' && line?.kind !== 'solid') {
		return current;
	}
	const next: PptxChartShapeProps = { ...current };
	if (fill?.kind === 'solid' && fill.color) {
		next.fillColor = withOpacity(safeColor(fill.color), opacityOf(fill.transparency));
	} else if (fill?.kind === 'none') {
		next.fillColor = 'none';
	}
	if (line?.kind === 'solid' && line.color) {
		next.strokeColor = withOpacity(safeColor(line.color), opacityOf(line.transparency));
		if (line.width !== undefined) next.strokeWidth = line.width;
	} else if (line?.kind === 'none') {
		next.strokeColor = 'none';
	}
	return next;
}

/**
 * Render-only copy of `data` whose data-label boxes follow the data-label
 * formats (colour, transparency, border), for the label decorations.
 */
export function withElementFormatLabelBoxes(data: PptxChartData): PptxChartData {
	const formats = data.elementFormats;
	if (!formats) {
		return data;
	}
	const chartWide = formats.dataLabels;
	const withShape = (
		opts: PptxChartDataLabelOptions | undefined,
		format: PptxChartElementFormat | undefined,
	) => {
		const shape = labelShapeFrom(format, opts?.labelShape);
		return shape === opts?.labelShape ? opts : { ...opts, labelShape: shape };
	};
	const style = data.style?.dataLabels
		? { ...data.style, dataLabels: withShape(data.style.dataLabels, chartWide) }
		: data.style;
	let changed = style !== data.style;
	const series = data.series.map((s, i) => {
		const format = formats[`dataLabels:${i}`] ?? chartWide;
		if (!s.dataLabelOptions || !format) return s;
		const opts = withShape(s.dataLabelOptions, format);
		if (opts === s.dataLabelOptions) return s;
		changed = true;
		return { ...s, dataLabelOptions: opts };
	});
	return changed ? { ...data, ...(style ? { style } : {}), series } : data;
}

/**
 * Shadow and glow of the data labels: cast by each label's box when the
 * labels have a fill or border, by the label text otherwise (that case is
 * drawn with the Text Options pass, from the label format's effects).
 */
export function dataLabelEffectsMarkup(vm: ChartViewModel, data: PptxChartData): string {
	const format = data.elementFormats?.dataLabels;
	const effects = format && !format.reset ? format.effects : undefined;
	if (!effects || (!effects.shadow && !(effects.glow && effects.glow.radius > 0))) return '';
	const boxed = paints(format?.fill, true) || format?.line?.kind === 'solid';
	if (!boxed || vm.dataLabels.length === 0) return '';
	const shapes = vm.dataLabels.map((label) => {
		const box = chartTextBox(label, 2);
		return {
			kind: 'rect' as const,
			x: box.left,
			y: box.top,
			w: box.right - box.left,
			h: box.bottom - box.top,
		};
	});
	return effectsBehind(shapes, effects);
}

/** Text-shaped shadow and glow of unboxed data labels, through the Text Options pass. */
export function withLabelTextEffects(data: PptxChartData): PptxChartData {
	const format = data.elementFormats?.dataLabels;
	const effects = format && !format.reset ? format.effects : undefined;
	if (!effects || paints(format?.fill, true) || format?.line?.kind === 'solid') return data;
	if (!effects.shadow && !(effects.glow && effects.glow.radius > 0)) return data;
	const text = format?.text ?? {};
	return {
		...data,
		elementFormats: {
			...data.elementFormats,
			dataLabels: {
				...format,
				text: {
					...text,
					effects: {
						...(effects.shadow ? { shadow: effects.shadow } : {}),
						...(effects.glow ? { glow: effects.glow } : {}),
						...text.effects,
					},
				},
			},
		},
	};
}

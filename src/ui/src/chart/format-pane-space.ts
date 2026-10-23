/**
 * A chart Format pane source over a `ChartSpace` (`ooxml-core/chart`): the
 * neutral view every product can give the pane once its chart model is a
 * `ChartSpace`. It offers what the neutral edit operations write (fill, line,
 * series, axis, legend and label options) and leaves out effects, 3-D, text,
 * pattern and picture fills and the secondary axis, which have no neutral
 * operation yet. Colours resolve through `theme` when given; a theme colour
 * without one shows as unset.
 *
 * @module chart/format-pane-space
 */
import type {
	ChartAxis,
	ChartFormatTarget,
	ChartPlotGroup,
	ChartSpace,
	ChartSpaceSeries,
} from 'ooxml-core/chart';
import {
	chartGroupPaintsLine,
	chartSeriesEntries,
	chartTargetShape,
	resolveChartGradient,
} from 'ooxml-core/chart';
import type {
	DrawingColor,
	DrawingColorTheme,
	DrawingFill,
	DrawingLine,
} from 'ooxml-core/drawingml';
import { resolveDrawingColor } from 'ooxml-core/drawingml';

import { chartFormatTargetOptions, type ChartFormatDescribe } from './format-pane-targets';
import type {
	ChartFormatAxisView,
	ChartFormatElementView,
	ChartFormatFillView,
	ChartFormatLabelsView,
	ChartFormatLineView,
	ChartFormatSeriesOptionsView,
	ChartFormatSource,
} from './format-pane-view';

export interface ChartSpaceFormatSourceOptions {
	/** Resolves scheme colours (`accent1`, `dk1`...) to hex. */
	theme?: DrawingColorTheme;
}

type Resolve = (color: DrawingColor) => { hex: string; alpha: number } | undefined;

const transparency = (alpha: number) => Math.round((1 - alpha) * 100);

function fillView(fill: DrawingFill | undefined, resolve: Resolve): ChartFormatFillView {
	switch (fill?.kind) {
		case 'none':
			return { kind: 'none' };
		case 'solid': {
			const color = resolve(fill.color);
			return color
				? { kind: 'solid', color: color.hex, transparency: transparency(color.alpha) }
				: { kind: 'solid' };
		}
		case 'gradient':
			return { kind: 'gradient', gradient: resolveChartGradient(fill, resolve) };
		case 'pattern':
			return { kind: 'pattern' };
		case 'picture':
			return { kind: 'picture' };
		default:
			return { kind: 'automatic' };
	}
}

function lineView(line: DrawingLine | undefined, resolve: Resolve): ChartFormatLineView {
	const paint = fillView(line?.fill, resolve);
	const kind = paint.kind === 'pattern' || paint.kind === 'picture' ? 'automatic' : paint.kind;
	return {
		kind,
		...(paint.color === undefined ? {} : { color: paint.color }),
		...(paint.transparency === undefined ? {} : { transparency: paint.transparency }),
		...(paint.gradient === undefined ? {} : { gradient: paint.gradient }),
		...(line?.widthEmu === undefined
			? {}
			: { width: Math.round((line.widthEmu / 12700) * 100) / 100 }),
		...(line?.dash === undefined ? {} : { dash: line.dash }),
		...(line?.cap === undefined ? {} : { cap: line.cap }),
	};
}

/** PowerPoint's name of a group's chart type (`barChart` is `bar`, `pie3DChart` is `pie3D`). */
const chartTypeOf = (group: ChartPlotGroup): string => group.element.replace(/Chart$/u, '');

const vertical = (axis: ChartAxis, space: ChartSpace): boolean => {
	if (axis.position) return axis.position === 'l' || axis.position === 'r';
	const horizontalBars = space.plotArea.groups.some((group) => group.barDirection === 'bar');
	return axis.kind === 'val' ? !horizontalBars : horizontalBars;
};

const seriesName = (series: ChartSpaceSeries | undefined, index: number) =>
	series?.tx?.text || series?.tx?.value || `Series ${index + 1}`;

function axisView(axis: ChartAxis): ChartFormatAxisView {
	const view: ChartFormatAxisView = {
		kind: axis.kind,
		reverse: axis.scaling.orientation === 'maxMin',
		logScale: axis.scaling.logBase !== undefined,
		numberFormat:
			axis.numberFormat && !axis.numberFormat.sourceLinked ? axis.numberFormat.formatCode : '',
	};
	const copy = <K extends keyof ChartFormatAxisView>(
		key: K,
		value: ChartFormatAxisView[K] | undefined,
	) => {
		if (value !== undefined) view[key] = value;
	};
	copy('min', axis.scaling.min);
	copy('max', axis.scaling.max);
	copy('logBase', axis.scaling.logBase);
	copy('majorUnit', axis.majorUnit);
	copy('minorUnit', axis.minorUnit);
	copy('crosses', axis.crosses);
	copy('crossesAt', axis.crossesAt);
	copy('displayUnits', axis.displayUnits?.builtInUnit);
	copy('majorTickMark', axis.majorTickMark);
	copy('minorTickMark', axis.minorTickMark);
	copy('tickLabelPosition', axis.tickLabelPosition);
	return view;
}

/** Every element of `space` a pane can target, in PowerPoint's picker order. */
function targetsOf(space: ChartSpace): ChartFormatTarget[] {
	const targets: ChartFormatTarget[] = [{ kind: 'chartArea' }];
	if (space.title) targets.push({ kind: 'title' });
	space.plotArea.axes.forEach((axis, axisIndex) => {
		if (axis.deleted) return;
		targets.push({ kind: 'axis', axisIndex });
		if (axis.majorGridlines) targets.push({ kind: 'majorGridlines', axisIndex });
		if (axis.minorGridlines) targets.push({ kind: 'minorGridlines', axisIndex });
	});
	if (space.legend) targets.push({ kind: 'legend' });
	targets.push({ kind: 'plotArea' });
	if (space.plotArea.groups.some((group) => group.dataLabels && !group.dataLabels.deleted))
		targets.push({ kind: 'dataLabels' });
	chartSeriesEntries(space).forEach((_, seriesIndex) => {
		targets.push({ kind: 'series', seriesIndex }, { kind: 'dataLabels', seriesIndex });
	});
	return targets;
}

/** The Format pane source of `space`. */
export function chartSpaceFormatSource(
	space: ChartSpace,
	options: ChartSpaceFormatSourceOptions = {},
): ChartFormatSource {
	const resolve: Resolve = (color) => {
		const resolved = resolveDrawingColor(color, options.theme);
		return resolved && { hex: resolved.hex, alpha: resolved.alpha };
	};
	const entries = chartSeriesEntries(space);
	const describe: ChartFormatDescribe = {
		axis: (axisIndex) => {
			const axis = space.plotArea.axes[axisIndex];
			if (!axis) return undefined;
			const secondary = space.plotArea.axes
				.slice(0, axisIndex)
				.some((other) => other.kind === axis.kind);
			return { kind: axis.kind, vertical: vertical(axis, space), secondary };
		},
		seriesName: (seriesIndex) => seriesName(entries[seriesIndex]?.series, seriesIndex),
	};
	const element = (target: ChartFormatTarget): ChartFormatElementView => {
		const lineTarget =
			target.kind === 'axis' ||
			target.kind === 'majorGridlines' ||
			target.kind === 'minorGridlines' ||
			(target.kind === 'series' &&
				entries[target.seriesIndex] !== undefined &&
				chartGroupPaintsLine(entries[target.seriesIndex]!.group));
		let shape;
		try {
			shape = chartTargetShape(space, target);
		} catch {
			shape = undefined;
		}
		return {
			support: {
				fill: !lineTarget,
				gradient: true,
				pattern: false,
				picture: false,
				threeD: false,
				text: false,
				lineDetails: false,
				effects: false,
			},
			fill: fillView(shape?.fill, resolve),
			line: lineView(shape?.line, resolve),
		};
	};
	const labels = (seriesIndex: number | undefined): ChartFormatLabelsView | undefined => {
		const entry = seriesIndex === undefined ? undefined : entries[seriesIndex];
		if (seriesIndex !== undefined && !entry) return undefined;
		const group = entry?.group ?? space.plotArea.groups[0];
		const chart = group?.dataLabels;
		const own = entry?.series.dataLabels;
		const pick = <K extends keyof NonNullable<typeof chart>>(key: K) => own?.[key] ?? chart?.[key];
		const view: ChartFormatLabelsView = { seriesScoped: entry !== undefined };
		const set = <K extends keyof ChartFormatLabelsView>(
			key: K,
			value: ChartFormatLabelsView[K] | undefined,
		) => {
			if (value !== undefined) view[key] = value;
		};
		set('chartType', group && chartTypeOf(group));
		set('showSeriesName', pick('showSeriesName'));
		set('showCategory', pick('showCategoryName'));
		set('showValue', pick('showValue'));
		set('showPercent', pick('showPercent'));
		set('showLegendKey', pick('showLegendKey'));
		set('position', pick('position'));
		return view;
	};
	return {
		targets: chartFormatTargetOptions(targetsOf(space), describe),
		textOptions: false,
		element,
		seriesOptions: (seriesIndex) => {
			const group = entries[seriesIndex]?.group;
			if (!group) return undefined;
			const view: ChartFormatSeriesOptionsView = { chartType: chartTypeOf(group) };
			if (group.overlap !== undefined) view.overlap = group.overlap;
			if (group.gapWidth !== undefined) view.gapWidth = group.gapWidth;
			if (group.firstSliceAngle !== undefined) view.firstSliceAngle = group.firstSliceAngle;
			if (group.holeSize !== undefined) view.holeSize = group.holeSize;
			return view;
		},
		axisOptions: (axisIndex) => {
			const axis = space.plotArea.axes[axisIndex];
			return axis && axisView(axis);
		},
		legendOptions: () =>
			space.legend && {
				...(space.legend.position === undefined ? {} : { position: space.legend.position }),
				overlay: space.legend.overlay ?? false,
			},
		labelOptions: labels,
	};
}

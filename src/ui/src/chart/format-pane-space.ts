/**
 * A chart Format pane source over a `ChartSpace` (`ooxml-core/chart`): the
 * neutral view every product can give the pane once its chart model is a
 * `ChartSpace`. It offers shared fill, outline, effects, text, 3-D, series
 * analysis, axis, legend and label edits. Each series' capabilities decide
 * which analysis and secondary-axis controls appear. Picture allocation stays
 * with the host. Colours resolve through `theme` when given.
 *
 * @module chart/format-pane-space
 */
import type { ChartAxis, ChartFormatTarget, ChartSpace, ChartSpaceSeries } from 'ooxml-core/chart';
import {
	chartGroupPaintsLine,
	chartSeriesEntries,
	chartTargetShape,
	resolveChartGradient,
	chartTargetHasText,
	chartSeriesSecondaryAxisSupported,
} from 'ooxml-core/chart';
import {
	chartAdvancedElementView,
	chartSeriesOptionsView,
	chartGroupFormatType as chartTypeOf,
} from './format-pane-advanced-space';
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
	/** Hosts with a narrower edit contract (for example ChartEx) can retain the basic pane. */
	advanced?: boolean;
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
			return {
				kind: 'pattern',
				pattern: {
					preset: fill.preset,
					foreground: fill.foreground ? (resolve(fill.foreground)?.hex ?? '#000000') : '#000000',
					background: fill.background ? (resolve(fill.background)?.hex ?? '#FFFFFF') : '#FFFFFF',
				},
			};
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
		...(line?.compound === undefined ? {} : { compound: line.compound }),
		...(line?.join === undefined ? {} : { join: line.join.kind }),
		...(line?.headEnd?.type === undefined ? {} : { beginArrow: line.headEnd.type }),
		...(line?.tailEnd?.type === undefined ? {} : { endArrow: line.tailEnd.type }),
	};
}

const vertical = (axis: ChartAxis, space: ChartSpace): boolean => {
	if (axis.position) return axis.position === 'l' || axis.position === 'r';
	const horizontalBars = space.plotArea.groups.some((group) => group.barDirection === 'bar');
	return axis.kind === 'val' ? !horizontalBars : horizontalBars;
};

const seriesName = (series: ChartSpaceSeries | undefined, index: number) =>
	series?.tx?.text || series?.tx?.value || `Series ${index + 1}`;

function axisView(axis: ChartAxis, advanced = false): ChartFormatAxisView {
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
	if (advanced) {
		view.title = axis.title?.text ?? '';
		if (axis.numberFormat) view.numberFormatSourceLinked = axis.numberFormat.sourceLinked ?? true;
		if (axis.kind === 'date') {
			view.timeUnitsEditable = true;
			copy('baseTimeUnit', axis.baseTimeUnit);
			copy('majorTimeUnit', axis.majorTimeUnit);
			copy('minorTimeUnit', axis.minorTimeUnit);
		}
	}
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
	const advanced = options.advanced !== false;
	const seriesOptions = (seriesIndex: number): ChartFormatSeriesOptionsView | undefined =>
		chartSeriesOptionsView(space, entries[seriesIndex], seriesIndex, advanced);
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
			...(target.kind === 'series' && seriesOptions(target.seriesIndex)
				? { series: seriesOptions(target.seriesIndex)! }
				: {}),
			...(advanced
				? chartAdvancedElementView(
						space,
						target,
						resolve,
						(fill) => fillView(fill, resolve),
						(line) => lineView(line, resolve),
						!lineTarget,
					)
				: {}),
			support: {
				fill: !lineTarget,
				gradient: true,
				pattern: advanced && !lineTarget,
				picture: false,
				threeD: advanced && !lineTarget,
				text: advanced && chartTargetHasText(target),
				lineDetails: advanced,
				effects: advanced,
				secondaryAxis:
					target.kind === 'series' &&
					advanced &&
					chartSeriesSecondaryAxisSupported(space, target.seriesIndex),
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
		textOptions: advanced,
		patternPresets: [
			'pct5',
			'pct10',
			'pct25',
			'pct50',
			'horz',
			'vert',
			'ltHorz',
			'cross',
			'diagCross',
		].map((preset) => ({
			value: preset,
			label: { key: `chartFormat.pattern.${preset}`, fallback: preset },
		})),
		element,
		seriesOptions,
		axisOptions: (axisIndex) => {
			const axis = space.plotArea.axes[axisIndex];
			return axis && axisView(axis, advanced);
		},
		legendOptions: () =>
			space.legend && {
				...(space.legend.position === undefined ? {} : { position: space.legend.position }),
				overlay: space.legend.overlay ?? false,
			},
		labelOptions: labels,
	};
}

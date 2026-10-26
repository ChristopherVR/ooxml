import {
	chartTargetEffects,
	chartTargetHasText,
	chartTargetTextFormatting,
	chartTargetThreeD,
	type ChartSpace,
	type ChartFormatTarget,
	type ChartSeriesEntry,
	type ChartPlotGroup,
	chartSeriesAnalysisSupported,
	chartSeriesMarkersSupported,
	chartSeriesSecondaryAxisSupported,
	chartSeriesOnSecondaryAxis,
} from 'ooxml-core/chart';
import type {
	DrawingColor,
	DrawingEffectList,
	DrawingFill,
	DrawingLine,
} from 'ooxml-core/drawingml';
import type {
	ChartFormatElementView,
	ChartFormatEffectsView,
	ChartFormatFillView,
	ChartFormatLineView,
	ChartFormatTextView,
	ChartFormatThreeDView,
	ChartFormatSeriesOptionsView,
} from './format-pane-view';

type Resolve = (color: DrawingColor) => { hex: string; alpha: number } | undefined;
const transparency = (alpha: number) => Math.round((1 - alpha) * 100);

export const chartGroupFormatType = (group: ChartPlotGroup): string =>
	group.element.replace(/Chart$/u, '');

export function chartSeriesOptionsView(
	space: ChartSpace,
	entry: ChartSeriesEntry | undefined,
	seriesIndex: number,
	advanced: boolean,
): ChartFormatSeriesOptionsView | undefined {
	const group = entry?.group;
	if (!entry || !group) return undefined;
	const view: ChartFormatSeriesOptionsView = { chartType: chartGroupFormatType(group) };
	for (const key of ['overlap', 'gapWidth', 'firstSliceAngle', 'holeSize'] as const)
		if (group[key] !== undefined) view[key] = group[key];
	if (advanced && chartSeriesSecondaryAxisSupported(space, seriesIndex))
		view.secondaryAxis = chartSeriesOnSecondaryAxis(space, seriesIndex);
	if (advanced && chartSeriesMarkersSupported(space, seriesIndex))
		view.marker = { ...entry.series.marker };
	if (advanced && chartSeriesAnalysisSupported(space, seriesIndex)) {
		const trends = entry.series.trendlines ?? [];
		const errors = (entry.series.errorBars ?? []).filter(
			(error) => (error.direction ?? 'y') === 'y',
		);
		view.analysis = {
			multipleTrendlines: trends.length > 1,
			multipleErrorBars: errors.length > 1,
			...(trends[0] ? { trendline: structuredClone(trends[0]) } : {}),
			...(errors[0] ? { errorBars: structuredClone(errors[0]) } : {}),
		};
	}
	return view;
}

export function chartEffectsView(
	list: DrawingEffectList | undefined,
	resolve: Resolve,
): ChartFormatEffectsView {
	const view: ChartFormatEffectsView = { shadow: null, glow: null, softEdge: 0 };
	for (const effect of list?.effects ?? []) {
		if (effect.kind === 'outerShadow') {
			const color = effect.color && resolve(effect.color);
			view.shadow = {
				color: color?.hex ?? '#000000',
				transparency: transparency(color?.alpha ?? 1),
				blur: (effect.blurRadiusEmu ?? 0) / 12700,
				distance: (effect.distanceEmu ?? 0) / 12700,
				angle: (effect.directionAngle ?? 0) / 60000,
			};
		} else if (effect.kind === 'glow') {
			const color = effect.color && resolve(effect.color);
			view.glow = {
				color: color?.hex ?? '#4472C4',
				transparency: transparency(color?.alpha ?? 1),
				radius: (effect.radiusEmu ?? 0) / 12700,
			};
		} else if (effect.kind === 'softEdge') view.softEdge = (effect.radiusEmu ?? 0) / 12700;
	}
	return view;
}

export function chartThreeDView(
	space: ChartSpace,
	target: ChartFormatTarget,
	resolve: Resolve,
): ChartFormatThreeDView {
	const properties = chartTargetThreeD(space, target);
	const color = (value: DrawingColor | null | undefined) =>
		value ? resolve(value)?.hex : undefined;
	const view: ChartFormatThreeDView = {};
	if (properties.bevelTop)
		view.bevelTop = {
			type: properties.bevelTop.preset,
			width: properties.bevelTop.widthEmu / 12700,
			height: properties.bevelTop.heightEmu / 12700,
		};
	if (properties.bevelBottom)
		view.bevelBottom = {
			type: properties.bevelBottom.preset,
			width: properties.bevelBottom.widthEmu / 12700,
			height: properties.bevelBottom.heightEmu / 12700,
		};
	const depth = color(properties.depthColor);
	const contour = color(properties.contourColor);
	view.depth = { size: (properties.depthEmu ?? 0) / 12700, ...(depth ? { color: depth } : {}) };
	view.contour = {
		size: (properties.contourWidthEmu ?? 0) / 12700,
		...(contour ? { color: contour } : {}),
	};
	if (properties.material !== undefined) view.material = properties.material;
	if (properties.lighting !== undefined) view.lighting = properties.lighting;
	return view;
}

export function chartAdvancedElementView(
	space: ChartSpace,
	target: ChartFormatTarget,
	resolve: Resolve,
	fillView: (fill: DrawingFill | undefined) => ChartFormatFillView,
	lineView: (line: DrawingLine | undefined) => ChartFormatLineView,
	threeD: boolean,
): Pick<ChartFormatElementView, 'effects' | 'text' | 'threeD'> {
	const result: Pick<ChartFormatElementView, 'effects' | 'text' | 'threeD'> = {
		effects: chartEffectsView(chartTargetEffects(space, target), resolve),
	};
	if (threeD) result.threeD = chartThreeDView(space, target, resolve);
	if (chartTargetHasText(target)) {
		const properties = chartTargetTextFormatting(space, target);
		const fill = fillView(properties.fill ?? undefined);
		const line = lineView(properties.line ?? undefined);
		const text: ChartFormatTextView = {
			fill: {
				...fill,
				kind: fill.kind === 'picture' || fill.kind === 'pattern' ? 'automatic' : fill.kind,
			},
			outline: { ...line, kind: line.kind === 'gradient' ? 'automatic' : line.kind },
			effects: chartEffectsView(properties.effects, resolve),
		};
		if (properties.rotation !== undefined) text.rotation = properties.rotation;
		if (properties.direction !== undefined) text.direction = properties.direction;
		if (properties.anchor !== undefined) text.anchor = properties.anchor;
		result.text = text;
	}
	return result;
}

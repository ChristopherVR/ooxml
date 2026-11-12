/**
 * The group settings a new classic chart part starts with (`chart-new-part.ts`):
 * the values the legacy generator wrote for the model, resolved here so both
 * writers agree. Gap width, overlap, first slice angle and hole size come from
 * `generatedChartGroupOptions` (the model's value rounded and clamped into its
 * schema range, else gap width 150 and hole size 50); a surface is a wireframe
 * unless the model says `wireframe: false`; a stock chart draws high-low lines
 * unless the model removes them (`hiLowLines: null`). The model's style,
 * colour variation, bar shape, gap depth, pie-of-pie and bubble options are
 * written as the model states them.
 *
 * @module utils/chart-new-part-group
 */
import {
	chartGroupElement,
	type ChartTypeTarget,
	type NewChartGroupSettings,
} from '../../../chart/index';
import type { PptxChartData } from '../types';
import { generatedChartGroupOptions } from './chart-group-options';
import { chartPercentUnionValue } from './chart-percent-union-value';

const OPTION_KEYS = {
	gapWidth: 'gapWidth',
	overlap: 'overlap',
	firstSliceAng: 'firstSliceAngle',
	holeSize: 'holeSize',
} as const;

/** The element-allowed settings of the new group, `undefined` members left out. */
export function newChartGroupSettings(
	chart: PptxChartData,
	target: ChartTypeTarget,
): NewChartGroupSettings {
	const element = chartGroupElement(target.kind, target.is3D ?? false);
	const settings: NewChartGroupSettings = {};
	for (const [option, value] of generatedChartGroupOptions(element, chart))
		settings[OPTION_KEYS[option]] = Number(value);
	const put = <K extends keyof NewChartGroupSettings>(
		key: K,
		value: NewChartGroupSettings[K] | undefined,
	) => {
		if (value !== undefined) settings[key] = value;
	};
	const varies = !['stockChart', 'surfaceChart', 'surface3DChart'].includes(element);
	if (varies) put('varyColors', chart.varyColors);
	if (element === 'bar3DChart') put('shape', chart.barShape);
	if (['bar3DChart', 'line3DChart', 'area3DChart'].includes(element))
		put('gapDepth', chart.gapDepth);
	if (target.kind === 'radar') put('radarStyle', chart.radarStyle);
	if (target.kind === 'scatter') put('scatterStyle', chart.scatterStyle);
	if (target.kind === 'surface') put('wireframe', chart.wireframe !== false);
	if (target.kind === 'stock' && chart.hiLowLines !== null) put('hiLowLines', true);
	if (target.kind === 'ofPie') {
		const options = chart.ofPieOptions;
		put('ofPieType', options?.ofPieType);
		put('splitType', options?.splitType);
		put('splitPosition', options?.splitPos);
		put('customSplit', options?.custSplit ? [...options.custSplit] : undefined);
		put('secondPieSize', options?.secondPieSize);
		if (options?.serLines) put('seriesLines', true);
	}
	if (target.kind === 'bubble') {
		const options = chart.bubbleOptions;
		put('bubble3D', options?.bubble3D);
		put(
			'bubbleScale',
			options?.bubbleScale === undefined
				? undefined
				: Number(
						chartPercentUnionValue(options.bubbleScale, { name: 'bubbleScale', min: 0, max: 300 }),
					),
		);
		put('showNegativeBubbles', options?.showNegativeBubbles);
		put('sizeRepresents', options?.sizeRepresents);
	}
	return settings;
}

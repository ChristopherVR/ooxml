/**
 * The six native charts of the deck, as SDK chart inputs. Each carries an embedded workbook, the
 * theme fonts, muted axes and gridlines, and exactly one highlighted series or point in coral.
 */
import type { ChartInput } from '../../../../src/core/pptx/core/builders/sdk/index';
import type { PptxChartAxisFormatting } from '../../../../src/core/pptx/core/types/chart';
import type { PptxElement } from '../../../../src/core/pptx/core/types/elements';
import { round } from '../../lib/random';
import {
	DISTRICT_MODE,
	DISTRICTS,
	FARES,
	FREQUENT_ROUTES,
	MODES,
	MONTHLY_2019,
	MONTHLY_2025,
	MONTHS,
	OTHER_ROUTES,
	PRODUCTIVITY,
	QUARTERS,
	SERVICE_HOURS,
	WEEKLY,
	sum,
} from './data';
import { C, CHART_FONT, MODE_COLORS } from './style';

const GRID_LINE = { color: '#E7E1D3', width: 1 };
const CATEGORY_AXIS = {
	font: CHART_FONT,
	majorTickMark: 'none',
	minorTickMark: 'none',
	line: { color: C.sandDeep, width: 1 },
} as const;
const VALUE_AXIS = {
	font: CHART_FONT,
	majorTickMark: 'none',
	minorTickMark: 'none',
	line: 'none',
	majorGridlines: true,
	majorGridlinesLine: GRID_LINE,
} as const;
const BARE = {
	embedWorkbook: true,
	hasLegend: false,
	chartArea: { fill: null, border: null },
	plotArea: { fill: null, border: null },
} as const;

/** Monthly boardings, 2025 against 2019; the record month is the one coral column. */
export function monthlyChart(): ChartInput {
	const record = MONTHLY_2025.indexOf(Math.max(...MONTHLY_2025));
	return {
		...BARE,
		categories: [...MONTHS],
		barDirection: 'col',
		grouping: 'clustered',
		gapWidth: 70,
		overlap: -8,
		series: [
			{ name: '2019', values: MONTHLY_2019, color: C.sandDeep },
			{
				name: '2025',
				values: MONTHLY_2025,
				color: C.teal,
				pointColors: MONTHLY_2025.map((_, index) => (index === record ? C.coral : null)),
			},
		],
		axes: {
			catAx: CATEGORY_AXIS,
			valAx: { ...VALUE_AXIS, min: 0, max: 5, majorUnit: 1, numberFormat: '0' },
		},
	};
}

/** The first week of each month carries the month's name; the other weeks are unlabelled. */
function weekLabels(): string[] {
	const starts = [1, 5, 9, 14, 18, 22, 27, 31, 36, 40, 44, 49];
	return Array.from({ length: 52 }, (_, index) => {
		const month = starts.indexOf(index + 1);
		return month >= 0 ? MONTHS[month]! : '';
	});
}

/** Weekly boardings by mode (thousands); the tram is the highlighted line. */
export function weeklyChart(): ChartInput {
	return {
		...BARE,
		categories: weekLabels(),
		series: MODES.map((mode, index) => ({
			name: mode,
			values: WEEKLY[index]!,
			color: MODE_COLORS[index]!,
			line: { width: mode === 'Tram' ? 3.5 : 2 },
			marker: { symbol: 'none' as const },
			smooth: false,
		})),
		axes: {
			catAx: CATEGORY_AXIS,
			valAx: { ...VALUE_AXIS, min: 0, max: 700, majorUnit: 100, numberFormat: '0' },
		},
	};
}

/** Mode share of each district's boardings, in percent (rows add up to 100). */
export function districtShares(): number[][] {
	return DISTRICT_MODE.map((row) => {
		const total = sum(row);
		return row.map((value) => round((value / total) * 100, 1));
	});
}

export function districtChart(): ChartInput {
	const shares = districtShares();
	return {
		...BARE,
		categories: DISTRICTS,
		barDirection: 'bar',
		grouping: 'stacked',
		gapWidth: 55,
		series: MODES.map((mode, index) => ({
			name: mode,
			values: shares.map((row) => row[index]!),
			color: MODE_COLORS[index]!,
			dataLabels: {
				showValue: true,
				position: 'ctr' as const,
				// Segments under 9% are too narrow for a label.
				numberFormat: '[<9]"";0"%"',
				font: { family: CHART_FONT.family, size: 10, bold: true, color: C.white },
			},
		})),
		axes: {
			catAx: {
				...CATEGORY_AXIS,
				orientation: 'maxMin',
				font: { ...CHART_FONT, size: 11, color: C.ink },
			},
			valAx: {
				...VALUE_AXIS,
				min: 0,
				max: 100,
				majorUnit: 25,
				numberFormat: '0"%"',
				tickLabelPosition: 'high',
			},
		},
	};
}

export const FARE_COLORS = [C.coral, C.teal, C.tealSoft, C.sandDeep, C.mist];

/** Fare products as a share of boardings. */
export function fareChart(): ChartInput {
	return {
		...BARE,
		categories: FARES.map((fare) => fare.name),
		holeSize: 66,
		firstSliceAngle: 0,
		series: [
			{
				name: 'Share of boardings',
				values: FARES.map((fare) => fare.share),
				pointColors: FARE_COLORS,
			},
		],
	};
}

/** Service hours as columns, boardings per service hour as a line on the secondary axis. */
export function productivityChart(): ChartInput {
	return {
		...BARE,
		categories: QUARTERS,
		gapWidth: 60,
		series: [
			{
				name: 'Service hours (thousand)',
				values: SERVICE_HOURS,
				chartType: 'bar',
				color: C.sandDeep,
			},
			{
				name: 'Boardings per service hour',
				values: PRODUCTIVITY,
				chartType: 'line',
				secondaryAxis: true,
				color: C.coral,
				line: { width: 3 },
				marker: { symbol: 'circle', size: 8, fill: C.coral, line: { color: C.paper, width: 1.5 } },
				dataLabels: {
					showValue: true,
					position: 't',
					numberFormat: '0.0',
					font: { family: CHART_FONT.family, size: 10, bold: true, color: C.coral },
				},
			},
		],
		axes: { catAx: CATEGORY_AXIS, valAx: { ...VALUE_AXIS, min: 0, max: 500, majorUnit: 100 } },
	};
}

/** Bus routes: peak frequency against boardings per service hour. */
export function frequencyChart(): ChartInput {
	const series = (name: string, points: typeof FREQUENT_ROUTES, color: string, size: number) => ({
		name,
		xValues: points.map((point) => point.frequency),
		values: points.map((point) => point.productivity),
		color,
		line: 'none' as const,
		marker: { symbol: 'circle' as const, size, fill: color, line: { color: C.paper, width: 1 } },
	});
	return {
		...BARE,
		categories: [],
		scatterStyle: 'marker',
		series: [
			series('Other bus routes', OTHER_ROUTES, C.tealSoft, 9),
			series('Frequent Network', FREQUENT_ROUTES, C.coral, 11),
		],
		axes: { valAx: CATEGORY_AXIS },
	};
}

/** The muted look of a value axis, on the chart model (for axes the SDK input cannot reach). */
const QUIET_AXIS: Partial<PptxChartAxisFormatting> = {
	fontFamily: CHART_FONT.family,
	fontSize: CHART_FONT.size,
	fontColor: CHART_FONT.color,
	majorTickMark: 'none',
	minorTickMark: 'none',
	spPr: { lineNoFill: true },
};

/**
 * The SDK's `axes.valAx` input formats one value axis only. The secondary axis of the combo
 * chart and the vertical axis of the scatter chart are formatted on the chart model instead.
 */
export function quietSecondAxis(
	chart: PptxElement,
	scale: { min: number; max: number; majorUnit: number },
): void {
	if (chart.type !== 'chart' || !chart.chartData) return;
	const axes = (chart.chartData.axes ??= []);
	const secondary = axes.find((axis) => axis.axisType === 'valAx' && axis.axPos === 'r');
	if (secondary) {
		Object.assign(secondary, QUIET_AXIS, scale);
		return;
	}
	// Scatter: the model lists only the horizontal axis; add the vertical one beside it.
	const horizontal = axes.find((axis) => axis.axisType === 'valAx');
	if (!horizontal) return;
	Object.assign(horizontal, {
		axisId: 111111111,
		crossAxisId: 222222222,
		axPos: 'b',
		min: 0,
		max: 14,
		majorUnit: 2,
	});
	axes.push({
		axisType: 'valAx',
		axisId: 222222222,
		crossAxisId: 111111111,
		axPos: 'l',
		...QUIET_AXIS,
		...scale,
		majorGridlines: true,
		majorGridlinesSpPr: { strokeColor: GRID_LINE.color, strokeWidth: GRID_LINE.width },
	});
}

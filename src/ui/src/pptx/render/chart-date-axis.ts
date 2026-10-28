import type { PptxChartAxisFormatting, PptxChartData } from 'ooxml-core/pptx';

import { chartAxisTextStyle, chartLineStyle, isNoLine } from './chart-axis-style';
import { formatDate } from './chart-date-format';
import type { PlotLayout, SvgLine, SvgText } from './chart-view-model';

import { excelSerialToDate, calendarTicks } from 'ooxml-core/chart';
export { excelSerialToDate } from 'ooxml-core/chart';

export interface DateAxisPlan {
	sourceIndices: number[];
	xPositions: number[];
	labels: SvgText[];
	tickMarks: SvgLine[];
}

function primaryDateAxis(chartData: PptxChartData): PptxChartAxisFormatting | undefined {
	return (
		chartData.axes?.find((axis) => axis.axisType === 'dateAx' && axis.axPos !== 't') ??
		chartData.axes?.find((axis) => axis.axisType === 'dateAx')
	);
}

function serialToX(serial: number, min: number, max: number, layout: PlotLayout): number {
	const ratio = max === min ? 0.5 : (serial - min) / (max - min);
	return layout.plotLeft + ratio * layout.plotWidth;
}

function tickLine(
	x: number,
	y: number,
	axis: PptxChartAxisFormatting,
	minor: boolean,
): SvgLine | undefined {
	const placement = minor ? axis.minorTickMark : axis.majorTickMark;
	if (!placement || placement === 'none' || isNoLine(axis.spPr)) {
		return undefined;
	}
	const length = minor ? 2.5 : 4;
	const direction = axis.axPos === 't' ? 1 : -1;
	const inward = direction * length;
	const outward = -inward;
	const [start, end] =
		placement === 'cross' ? [-length, length] : placement === 'in' ? [0, inward] : [0, outward];
	return {
		kind: 'line',
		x1: x,
		y1: y + start,
		x2: x,
		y2: y + end,
		...chartLineStyle(axis.spPr),
	};
}

/** Build a continuous calendar scale for classic ChartML date categories. */
export function buildDateAxisPlan(
	chartData: PptxChartData,
	layout: PlotLayout,
	selectedAxis?: PptxChartAxisFormatting,
	axisY?: number,
): DateAxisPlan | undefined {
	const dates = chartData.dateCategories;
	const axis = selectedAxis ?? primaryDateAxis(chartData);
	if (!dates || !axis || dates.values.length === 0) {
		return undefined;
	}
	const indexed = dates.values
		.map((value, index) => ({ value, index }))
		.sort((a, b) => a.value - b.value);
	const dataMin = indexed[0]?.value ?? 0;
	const dataMax = indexed[indexed.length - 1]?.value ?? dataMin;
	const min = axis.min ?? dataMin;
	const max = axis.max ?? dataMax;
	const visible = indexed.filter(({ value }) => value >= min && value <= max);
	if (axis.orientation === 'maxMin') {
		visible.reverse();
	}
	const sourceIndices = visible.map(({ index }) => index);
	const xPositions = visible.map(({ value }) => {
		const x = serialToX(value, min, max, layout);
		return axis.orientation === 'maxMin'
			? layout.plotLeft + layout.plotWidth - (x - layout.plotLeft)
			: x;
	});
	const unit = axis.majorTimeUnit ?? axis.baseTimeUnit ?? 'days';
	const ticks = calendarTicks(
		min,
		max,
		dates.date1904 ?? false,
		unit,
		Math.max(1, axis.majorUnit ?? 1),
	);
	const y = axisY ?? (axis.axPos === 't' ? layout.plotTop : layout.plotBottom);
	const offset = 4 + 8 * ((axis.labelOffset ?? 100) / 100);
	const labelsAbove =
		axis.tickLblPos === 'high' || (axis.tickLblPos !== 'low' && axis.axPos === 't');
	const labelY =
		axis.tickLblPos === 'high' ? layout.plotTop : axis.tickLblPos === 'low' ? layout.plotBottom : y;
	const labelSkip = Math.max(1, axis.tickLabelSkip ?? 1);
	const labels =
		axis.deleted || axis.tickLblPos === 'none'
			? []
			: ticks.flatMap((serial, index) => {
					if (index % labelSkip !== 0) {
						return [];
					}
					let x = serialToX(serial, min, max, layout);
					if (axis.orientation === 'maxMin') {
						x = layout.plotLeft + layout.plotWidth - (x - layout.plotLeft);
					}
					return [
						{
							kind: 'text' as const,
							x,
							y: labelsAbove ? labelY - offset : labelY + offset,
							text: formatDate(
								excelSerialToDate(serial, dates.date1904),
								unit,
								axis.numFmt?.formatCode ?? dates.formatCode,
							),
							...chartAxisTextStyle(axis),
							textAnchor: 'middle' as const,
						},
					];
				});
	const tickMarks = axis.deleted
		? []
		: ticks.flatMap((serial, index) => {
				if (index % Math.max(1, axis.tickMarkSkip ?? 1) !== 0) {
					return [];
				}
				let x = serialToX(serial, min, max, layout);
				if (axis.orientation === 'maxMin') {
					x = layout.plotLeft + layout.plotWidth - (x - layout.plotLeft);
				}
				const line = tickLine(x, y, axis, false);
				return line ? [line] : [];
			});
	if (!axis.deleted && axis.spPr && !isNoLine(axis.spPr)) {
		tickMarks.unshift({
			kind: 'line',
			x1: layout.plotLeft,
			y1: y,
			x2: layout.plotRight,
			y2: y,
			...chartLineStyle(axis.spPr),
			...(axis.axisId !== undefined ? { axisLine: axis.axisId } : {}),
		});
	}
	if (!axis.deleted && axis.minorTickMark && axis.minorTickMark !== 'none') {
		const minorUnit = axis.minorTimeUnit ?? axis.baseTimeUnit ?? unit;
		for (const serial of calendarTicks(
			min,
			max,
			dates.date1904 ?? false,
			minorUnit,
			Math.max(1, axis.minorUnit ?? 1),
		)) {
			let x = serialToX(serial, min, max, layout);
			if (axis.orientation === 'maxMin') {
				x = layout.plotLeft + layout.plotWidth - (x - layout.plotLeft);
			}
			const line = tickLine(x, y, axis, true);
			if (line) {
				tickMarks.push(line);
			}
		}
	}
	return { sourceIndices, xPositions, labels, tickMarks };
}

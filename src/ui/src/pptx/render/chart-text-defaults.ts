/**
 * chart-text-defaults.ts: resolves chart text styles field by field, starting
 * from the chart-wide style (`c:chartSpace/c:txPr` over the theme's minor
 * fonts, `PptxChartStyle.textStyle`).
 *
 * PowerPoint draws any chart text whose own `txPr` leaves a field unset with
 * the value from the next level up, ending at the chart-wide style. Core keeps
 * each level as authored, so a save never copies one level into another; the
 * cascade is applied here, on what the view-model is built from.
 *
 * @module chart-text-defaults
 */
import type { PptxChartData, PptxChartLegendTextStyle } from 'ooxml-core/pptx';

import { chartTextFontFamily } from './chart-font';

/**
 * Merge `levels` (outermost first) over the chart-wide text style, field by
 * field. Returns `undefined` when no level sets anything.
 */
export function resolveChartTextStyle(
	chartData: PptxChartData | undefined,
	...levels: ReadonlyArray<PptxChartLegendTextStyle | undefined>
): PptxChartLegendTextStyle | undefined {
	const style: PptxChartLegendTextStyle = {};
	for (const level of [chartData?.style?.textStyle, ...levels]) {
		for (const [key, value] of Object.entries(level ?? {})) {
			if (value !== undefined) {
				Object.assign(style, { [key]: value });
			}
		}
	}
	return Object.keys(style).length > 0 ? style : undefined;
}

/** `chartData` with each axis's text resolved through the chart-wide style. */
export function withChartTextDefaults(chartData: PptxChartData): PptxChartData {
	const axes = chartData.axes;
	if (!axes?.length || (!chartData.style?.textStyle && !axes.some((a) => a.eastAsiaFontFamily))) {
		return chartData;
	}
	return {
		...chartData,
		axes: axes.map((axis) => {
			const text = resolveChartTextStyle(chartData, {
				fontFamily: axis.fontFamily,
				eastAsiaFontFamily: axis.eastAsiaFontFamily,
				fontSize: axis.fontSize,
				bold: axis.fontBold,
				color: axis.fontColor,
			});
			const fontFamily = chartTextFontFamily(text);
			return {
				...axis,
				...(fontFamily ? { fontFamily } : {}),
				...(text?.fontSize !== undefined ? { fontSize: text.fontSize } : {}),
				...(text?.bold !== undefined ? { fontBold: text.bold } : {}),
				...(text?.color ? { fontColor: text.color } : {}),
			};
		}),
	};
}

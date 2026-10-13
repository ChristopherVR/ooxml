import type { ChartPptxElement } from '../../index';

const LITERAL_CATEGORY_TYPES = new Set([
	'bar',
	'line',
	'pie',
	'ofPie',
	'doughnut',
	'area',
	'radar',
	'bar3D',
	'line3D',
	'pie3D',
	'area3D',
	'surface',
]);

/** Unsaved SDK charts have no package sources: their labels are generated as literals. */
export function chartLabelsAreEditable(element: ChartPptxElement): boolean {
	const data = element.chartData;
	if (!data) return false;
	if (data.dataLabelTextEditable !== undefined) {
		const shape = data.dataLabelSourceShape;
		return (
			data.dataLabelTextEditable &&
			(!shape ||
				(shape.categoryCount === data.categories.length &&
					shape.seriesValueCounts.length === data.series.length &&
					shape.seriesValueCounts.every(
						(count, index) => count === data.series[index]?.values.length,
					)))
		);
	}
	return (
		!element.rawXml &&
		!data.chartPartPath &&
		!data.externalData &&
		!data.categoryLevels &&
		!data.dateCategories &&
		LITERAL_CATEGORY_TYPES.has(data.chartType) &&
		data.categories.length > 0 &&
		data.series.length > 0 &&
		data.series.every(
			(series) =>
				series.values.length === data.categories.length && series.values.every(Number.isFinite),
		)
	);
}

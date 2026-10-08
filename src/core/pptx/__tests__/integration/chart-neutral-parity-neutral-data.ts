// Chart parity harness: the neutral side's data helpers. Reads the neutral model's caches and text
// the way the pptx parser reads the object tree, so the two can be compared.
import type { ChartDataCache, ChartDataSource, ChartPlotGroup } from '../../../chart/index';
import type { DrawingTextBody } from '../../../drawingml/index';
import type { PptxChartType } from '../../core/types';
import { nonEmpty } from './chart-neutral-parity';

/** pptx chart type tokens by chart-group element (`detectChartType`). */
export const GROUP_TYPE: Record<string, PptxChartType> = {
	barChart: 'bar',
	bar3DChart: 'bar3D',
	lineChart: 'line',
	line3DChart: 'line3D',
	pieChart: 'pie',
	pie3DChart: 'pie3D',
	ofPieChart: 'ofPie',
	doughnutChart: 'doughnut',
	areaChart: 'area',
	area3DChart: 'area3D',
	scatterChart: 'scatter',
	bubbleChart: 'bubble',
	radarChart: 'radar',
	stockChart: 'stock',
	surfaceChart: 'surface',
	surface3DChart: 'surface',
};

/** Concatenated run text of a text body, without paragraph breaks (`collectAllText`). */
export function runText(body: DrawingTextBody | undefined): string | undefined {
	if (!body) return undefined;
	return body.paragraphs.flatMap((paragraph) => paragraph.runs.map((run) => run.text)).join('');
}

/**
 * The cache expanded to its declared length, keyed by `c:pt/@idx`; a missing point is undefined.
 * Both parsers expand sparse caches this way (`extractSeriesNumbersWithBlanks`,
 * `extractChartCategoryValues`).
 */
function expand(cache: ChartDataCache | undefined): (string | undefined)[] {
	if (!cache) return [];
	const byIndex = new Map<number, string>();
	let maxIndex = -1;
	for (const point of cache.points) {
		const index = point.index ?? 0;
		byIndex.set(index, point.value);
		maxIndex = Math.max(maxIndex, index);
	}
	const length = Math.max(cache.pointCount ?? 0, maxIndex + 1);
	if (byIndex.size === 0) return [];
	return Array.from({ length }, (_, index) => byIndex.get(index));
}

/** pptx normalisation: point text is trimmed, and a non-numeric point is a blank. */
export function numbers(source: ChartDataSource | undefined): (number | null)[] {
	if (source?.kind !== 'numRef' && source?.kind !== 'numLit') return [];
	return expand(source.cache).map((value) => {
		const text = value?.trim() ?? '';
		const number = Number.parseFloat(text);
		return text.length > 0 && Number.isFinite(number) ? number : null;
	});
}

/** pptx normalisation: category text is trimmed and a blank slot is the empty string. */
function strings(cache: ChartDataCache | undefined): string[] {
	const values = expand(cache).map((value) => value?.trim() ?? '');
	return values.some((value) => value.length > 0) ? values : [];
}

function categoryStrings(source: ChartDataSource | undefined, numeric: boolean): string[] {
	const kinds = numeric ? ['numRef', 'numLit'] : ['strRef', 'strLit'];
	return source?.kind && kinds.includes(source.kind) ? strings(source.cache) : [];
}

/**
 * The chart-level category list the pptx model keeps: the first series (of the first chart group
 * that yields any) supplies it; a multi-level source gives its innermost level, and the outer
 * levels are forward-filled across their spans (pptx normalisation: a group label is repeated on
 * every leaf it spans, `extractMultiLevelCategoryValues`).
 */
export function neutralCategories(groups: ChartPlotGroup[]): {
	categories: string[];
	categoryLevels?: string[][];
} {
	for (const group of groups) {
		const first = group.series[0];
		if (!first) continue;
		const source = first.categories;
		if (source?.kind === 'multiLvlStrRef' && source.levels && source.levels.length > 0) {
			const declared = source.cache?.pointCount;
			const levels = source.levels.map((level, levelIndex) => {
				const dense = expand({
					...level,
					...(declared !== undefined ? { pointCount: declared } : {}),
				}).map((value) => value?.trim() ?? '');
				if (levelIndex > 0) {
					let last = '';
					for (let index = 0; index < dense.length; index++) {
						if (dense[index]) last = dense[index];
						else dense[index] = last;
					}
				}
				return dense;
			});
			return {
				categories: levels[0] ?? [],
				...(levels.length > 1 ? { categoryLevels: levels } : {}),
			};
		}
		const categories =
			nonEmpty(categoryStrings(source, false)) ??
			nonEmpty(categoryStrings(source, true)) ??
			categoryStrings(first.xValues, false);
		if (categories.length > 0) return { categories };
	}
	return { categories: [] };
}

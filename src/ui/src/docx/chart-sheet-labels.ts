import type { OfficeChartDataSheetLabels } from '../chart/data-sheet';
import { translate, type EditorLocale } from './localization';

export function sheetLabels(locale: EditorLocale): OfficeChartDataSheetLabels {
	const t = (key: keyof OfficeChartDataSheetLabels) =>
		translate(locale, `chart.sheet.${key}` as Parameters<typeof translate>[1]);
	return {
		heading: t('heading'),
		addCategory: t('addCategory'),
		addSeries: t('addSeries'),
		addCategoryShort: t('addCategoryShort'),
		addSeriesShort: t('addSeriesShort'),
		removeCategory: t('removeCategory'),
		removeSeries: t('removeSeries'),
		categories: t('categories'),
		seriesName: t('seriesName'),
		categoryLabel: t('categoryLabel'),
		valueCell: t('valueCell'),
		summary: t('summary'),
		invalidNumber: t('invalidNumber'),
	};
}

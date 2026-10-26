import { normalizeLocale } from 'ooxml-core/i18n';
import { CHART_DATA_SHEET_ENGLISH, type OfficeChartDataSheetLabels } from './data-sheet-labels';
const keys = Object.keys(CHART_DATA_SHEET_ENGLISH) as (keyof OfficeChartDataSheetLabels)[];
const localized = {
	de: [
		'Daten',
		'Kategorie hinzufügen',
		'Datenreihe hinzufügen',
		'Kat.',
		'Datenreihe',
		'Kategorie entfernen',
		'Datenreihe entfernen',
		'Kategorien',
		'Datenreihe {number}',
		'Kategorien {number}',
		'{name}, Wert {number}',
		'{series} Datenreihen · {categories} Kategorien',
		'Eine Zahl eingeben',
	],
	fr: [
		'Données',
		'Ajouter une catégorie',
		'Ajouter une série',
		'Cat.',
		'Série',
		'Supprimer la catégorie',
		'Supprimer la série',
		'Catégories',
		'Série {number}',
		'Catégories {number}',
		'Valeur {number} de {name}',
		'{series} séries · {categories} catégories',
		'Saisir un nombre',
	],
	es: [
		'Datos',
		'Agregar categoría',
		'Agregar serie',
		'Cat.',
		'Serie',
		'Quitar categoría',
		'Quitar serie',
		'Categorías',
		'Serie {number}',
		'Categorías {number}',
		'Valor {number} de {name}',
		'{series} series · {categories} categorías',
		'Introduzca un número',
	],
	'zh-CN': [
		'数据',
		'添加类别',
		'添加系列',
		'类别',
		'系列',
		'删除类别',
		'删除系列',
		'类别',
		'系列 {number}',
		'类别 {number}',
		'{name} 的数值 {number}',
		'{series} 个系列 · {categories} 个类别',
		'请输入数字',
	],
} as const;
/** Captions for the shared chart data sheet, independent of product dictionaries. */
export function chartDataSheetLabels(locale: string): OfficeChartDataSheetLabels {
	const normalized = normalizeLocale(locale, ['en', 'de', 'fr', 'es', 'zh-CN'] as const, 'en');
	if (normalized === 'en') return CHART_DATA_SHEET_ENGLISH;
	return Object.fromEntries(
		keys.map((key, index) => [key, localized[normalized][index] ?? CHART_DATA_SHEET_ENGLISH[key]]),
	) as unknown as OfficeChartDataSheetLabels;
}

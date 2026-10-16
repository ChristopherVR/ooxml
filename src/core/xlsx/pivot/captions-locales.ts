import type { PivotAggregate, PivotCaptionOverrides } from '../model';

/**
 * The captions Excel writes in its other UI languages, for the languages the editors ship. Only
 * the English Excel is at hand, so these tables follow the wording of the localized Excel
 * releases rather than a measured file; a workbook that holds other words is read as it is.
 */
type Words = Record<PivotAggregate, string>;
const table = (
	words: Words,
	prefix: (word: string) => string,
	rest: Omit<PivotCaptionOverrides, 'functions' | 'aggregates'>,
	functions: Words,
): PivotCaptionOverrides => ({
	...rest,
	functions,
	aggregates: Object.fromEntries(
		(Object.keys(words) as PivotAggregate[]).map((fn) => [fn, prefix(words[fn])]),
	) as Words,
});

const GERMAN: Words = {
	sum: 'Summe',
	count: 'Anzahl',
	average: 'Mittelwert',
	max: 'Max',
	min: 'Min',
	product: 'Produkt',
	countNums: 'Anzahl (Zahlen)',
	stdDev: 'Standardabweichung',
	stdDevp: 'Standardabweichung (Grundgesamtheit)',
	var: 'Varianz',
	varp: 'Varianz (Grundgesamtheit)',
};
const SPANISH: Words = {
	sum: 'Suma',
	count: 'Cuenta',
	average: 'Promedio',
	max: 'Máx',
	min: 'Mín',
	product: 'Producto',
	countNums: 'Cuenta de números',
	stdDev: 'Desvest',
	stdDevp: 'Desvestp',
	var: 'Var',
	varp: 'Varp',
};
const FRENCH: Words = {
	sum: 'Somme',
	count: 'Nombre',
	average: 'Moyenne',
	max: 'Max',
	min: 'Min',
	product: 'Produit',
	countNums: 'Nombre (nombres)',
	stdDev: 'Écartype',
	stdDevp: 'Écartypep',
	var: 'Var',
	varp: 'Varp',
};
const CHINESE: Words = {
	sum: '求和',
	count: '计数',
	average: '平均值',
	max: '最大值',
	min: '最小值',
	product: '乘积',
	countNums: '数值计数',
	stdDev: '标准偏差',
	stdDevp: '总体标准偏差',
	var: '方差',
	varp: '总体方差',
};

export const LOCALIZED_PIVOT_CAPTIONS: Record<string, PivotCaptionOverrides> = {
	de: table(
		GERMAN,
		(word) => `${word} von {0}`,
		{
			rowLabels: 'Zeilenbeschriftungen',
			columnLabels: 'Spaltenbeschriftungen',
			grandTotal: 'Gesamtergebnis',
			values: 'Werte',
			all: '(Alle)',
			blank: '(leer)',
			subtotal: '{0} Ergebnis',
			measureTotal: 'Gesamt {0}',
		},
		GERMAN,
	),
	es: table(
		SPANISH,
		(word) => `${word} de {0}`,
		{
			rowLabels: 'Etiquetas de fila',
			columnLabels: 'Etiquetas de columna',
			grandTotal: 'Total general',
			values: 'Valores',
			all: '(Todas)',
			blank: '(en blanco)',
			subtotal: '{0} Total',
			measureTotal: 'Total {0}',
		},
		SPANISH,
	),
	fr: table(
		FRENCH,
		(word) => `${word} de {0}`,
		{
			rowLabels: 'Étiquettes de lignes',
			columnLabels: 'Étiquettes de colonnes',
			grandTotal: 'Total général',
			values: 'Valeurs',
			all: '(Tous)',
			blank: '(vide)',
			subtotal: '{0} Total',
			measureTotal: 'Total {0}',
		},
		FRENCH,
	),
	'zh-CN': table(
		CHINESE,
		(word) => `${word}项:{0}`,
		{
			rowLabels: '行标签',
			columnLabels: '列标签',
			grandTotal: '总计',
			values: '数值',
			all: '(全部)',
			blank: '(空白)',
			subtotal: '{0} 汇总',
			measureTotal: '总计 {0}',
		},
		CHINESE,
	),
};

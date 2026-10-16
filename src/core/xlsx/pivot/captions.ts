import { normalizeOfficeLocale } from '../../i18n/index';
import type { PivotAggregate, PivotCaptionOverrides, PivotCaptions, PivotTable } from '../model';
import { LOCALIZED_PIVOT_CAPTIONS } from './captions-locales';

const FUNCTIONS: Record<PivotAggregate, string> = {
	sum: 'Sum',
	count: 'Count',
	average: 'Average',
	max: 'Max',
	min: 'Min',
	product: 'Product',
	countNums: 'Count',
	stdDev: 'StdDev',
	stdDevp: 'StdDevp',
	var: 'Var',
	varp: 'Varp',
};
/** The default value field name of English Excel: `Sum of Amount`, `CountNums of Amount`. */
const aggregateNames = (): Record<PivotAggregate, string> =>
	Object.fromEntries(
		(Object.keys(FUNCTIONS) as PivotAggregate[]).map((fn) => [
			fn,
			`${fn[0]!.toUpperCase()}${fn.slice(1)} of {0}`,
		]),
	) as Record<PivotAggregate, string>;

/** What English Excel writes. */
export const ENGLISH_PIVOT_CAPTIONS: PivotCaptions = {
	rowLabels: 'Row Labels',
	columnLabels: 'Column Labels',
	grandTotal: 'Grand Total',
	values: 'Values',
	all: '(All)',
	blank: '(blank)',
	subtotal: '{0} Total',
	measureTotal: 'Total {0}',
	functions: FUNCTIONS,
	aggregates: aggregateNames(),
};

/** Anything that carries the captions of a report. */
export type CaptionHolder = { captions?: PivotCaptionOverrides | undefined };

/** A template with `{0}`, `{1}` replaced by the arguments. */
export const fillCaption = (template: string, ...args: string[]): string =>
	template.replace(/\{(\d)\}/g, (_match, at: string) => args[Number(at)] ?? '');

/** The captions of a report: its own, over English. */
export function pivotCaptions(pivot: CaptionHolder): PivotCaptions {
	const own = pivot.captions;
	if (!own) return ENGLISH_PIVOT_CAPTIONS;
	return {
		...ENGLISH_PIVOT_CAPTIONS,
		...Object.fromEntries(Object.entries(own).filter(([, value]) => typeof value === 'string')),
		functions: { ...ENGLISH_PIVOT_CAPTIONS.functions, ...own.functions },
		aggregates: { ...ENGLISH_PIVOT_CAPTIONS.aggregates, ...own.aggregates },
	};
}

/** The default name of a value field: `Sum of Amount`, `Summe von Amount`. */
export const pivotMeasureName = (
	aggregate: PivotAggregate,
	field: string,
	pivot: CaptionHolder = {},
): string => fillCaption(pivotCaptions(pivot).aggregates[aggregate], field);

/**
 * The captions Excel writes in the UI language of `locale` (`de`, `es-ES`, `zh-CN`): English
 * gives none. The editors pass their own locale when they author a report.
 */
export function localizedPivotCaptions(
	locale: string | undefined,
): PivotCaptionOverrides | undefined {
	const caps = LOCALIZED_PIVOT_CAPTIONS[normalizeOfficeLocale(locale)];
	return caps ? structuredClone(caps) : undefined;
}

/** Drops what equals English, so a report in English carries no captions at all. */
export function trimmedCaptions(caps: PivotCaptions): PivotCaptionOverrides | undefined {
	const out: PivotCaptionOverrides = {};
	for (const key of [
		'rowLabels',
		'columnLabels',
		'grandTotal',
		'values',
		'all',
		'blank',
		'subtotal',
		'measureTotal',
	] as const)
		if (caps[key] !== ENGLISH_PIVOT_CAPTIONS[key]) out[key] = caps[key];
	for (const key of ['functions', 'aggregates'] as const) {
		const changed = (Object.keys(caps[key]) as PivotAggregate[]).filter(
			(fn) => caps[key][fn] !== ENGLISH_PIVOT_CAPTIONS[key][fn],
		);
		if (changed.length) out[key] = Object.fromEntries(changed.map((fn) => [fn, caps[key][fn]]));
	}
	return Object.keys(out).length ? out : undefined;
}

/** The caption overrides a pivot edit keeps: the report's own, else the authoring ones. */
export function authoredCaptions(options: {
	captions?: PivotCaptionOverrides | undefined;
	locale?: string | undefined;
}): PivotTable['captions'] {
	const own = options.captions ?? localizedPivotCaptions(options.locale);
	return own && Object.keys(own).length ? structuredClone(own) : undefined;
}

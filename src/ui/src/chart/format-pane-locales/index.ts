/**
 * The shared chart Format pane's own translations: every `chartFormat.*` caption in each display
 * locale the products ship (English, French, German, Spanish, Simplified Chinese, Japanese,
 * Korean). Pure data with no imports, so a product dictionary can build on it
 * (`ooxml-ui/chart-format-locales`).
 *
 * @module chart/format-pane-locales
 */
import { chartFormatDe } from './de';
import { chartFormatEn } from './en';
import { chartFormatEs } from './es';
import { chartFormatFr } from './fr';
import { chartFormatJa } from './ja';
import { chartFormatKo } from './ko';
import { chartFormatZhCN } from './zh-CN';

export const CHART_FORMAT_LOCALES = ['en', 'fr', 'de', 'es', 'zh-CN', 'ja', 'ko'] as const;
export type ChartFormatLocale = (typeof CHART_FORMAT_LOCALES)[number];

/** `chartFormat.*` key -> caption, per locale. */
export const CHART_FORMAT_STRINGS: Readonly<
	Record<ChartFormatLocale, Readonly<Record<string, string>>>
> = {
	en: chartFormatEn,
	fr: chartFormatFr,
	de: chartFormatDe,
	es: chartFormatEs,
	'zh-CN': chartFormatZhCN,
	ja: chartFormatJa,
	ko: chartFormatKo,
};

/** `strings` with every key prefixed (PowerPoint's dictionary keeps them as `pptx.chartFormat.*`). */
export function prefixChartFormatStrings(
	strings: Readonly<Record<string, string>>,
	prefix: string,
): Record<string, string> {
	return Object.fromEntries(Object.entries(strings).map(([key, text]) => [prefix + key, text]));
}

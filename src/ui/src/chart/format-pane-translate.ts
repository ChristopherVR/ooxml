/**
 * The chart Format pane's translator: a product's own captions (`translateLabel`) first, then the
 * pane's shared dictionary for the display locale, then English.
 *
 * @module chart/format-pane-translate
 */
import { normalizeLocale } from 'ooxml-core/i18n';
import type { ChartFormatTranslate } from './format-pane-common';
import {
	CHART_FORMAT_LOCALES,
	CHART_FORMAT_STRINGS,
	type ChartFormatLocale,
} from './format-pane-locales/index';

/** A supported pane locale for any BCP 47 tag (`fr-CA` is French), English otherwise. */
export const chartFormatLocale = (value: string | null | undefined): ChartFormatLocale =>
	normalizeLocale(value, CHART_FORMAT_LOCALES, 'en');

/**
 * Translates `chartFormat.*` captions for `locale`. `override` (a product's translator) wins
 * where it knows the key, that is where it returns neither the key nor an empty string.
 */
export function chartFormatTranslator(
	locale: string | null | undefined,
	override?: ChartFormatTranslate,
): ChartFormatTranslate {
	const strings = CHART_FORMAT_STRINGS[chartFormatLocale(locale)];
	return (key, params, fallback) => {
		const own = override?.(key, params, fallback);
		if (own && own !== key) return own;
		return strings[key] ?? CHART_FORMAT_STRINGS.en[key] ?? key;
	};
}

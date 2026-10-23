/**
 * Caption and colour helpers every chart Format pane section shares. Captions
 * carry `chartFormat.*` keys with English fallbacks; a product's translator
 * maps them onto its own dictionary (PowerPoint prefixes `pptx.`, Excel looks
 * the English text up).
 *
 * @module chart/format-pane-common
 */
import type { ChartFormatChoice, ChartFormatText } from './format-pane-types';

export const t = (key: string, fallback: string): ChartFormatText => ({
	key: `chartFormat.${key}`,
	fallback,
});

export const choice = (
	value: string,
	key: string,
	fallback: string,
	extra?: Partial<ChartFormatChoice>,
): ChartFormatChoice => ({ value, label: t(key, fallback), ...extra });

/** Percent 0..100 for display. */
export const pct = (value: number | undefined): number => Math.round(value ?? 0);

/** `#RRGGBB` for an `<input type=color>`; unknown colours fall back to black. */
export function hexForInput(color: string | undefined, fallback = '#000000'): string {
	if (color && /^#?[0-9a-f]{6}$/iu.test(color)) {
		return `#${color.replace(/^#/u, '').toUpperCase()}`;
	}
	return fallback;
}

/**
 * A translator: the caption's key, its parameters and the English fallback. It returns the key
 * itself (or an empty string) when it has no translation, so the fallback is used.
 */
export type ChartFormatTranslate = (
	key: string,
	params?: Readonly<Record<string, string | number>>,
	fallback?: string,
) => string;

const fill = (text: string, params: Readonly<Record<string, string | number>> | undefined) =>
	text.replace(/\{\{(\w+)\}\}/gu, (_, name: string) => String(params?.[name] ?? ''));

/** Resolves a caption with `tr`, filling `{{name}}` parameters in either result. */
export function translateChartFormatText(value: ChartFormatText, tr: ChartFormatTranslate): string {
	const nested = Object.entries(value.paramTexts ?? {}).map(
		([name, caption]) => [name, translateChartFormatText(caption, tr)] as const,
	);
	const params =
		value.params || nested.length > 0
			? { ...value.params, ...Object.fromEntries(nested) }
			: undefined;
	const translated = tr(value.key, params, value.fallback);
	return fill(translated && translated !== value.key ? translated : value.fallback, params);
}

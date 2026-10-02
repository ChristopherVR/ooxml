/**
 * Display-language strings for the editor UI. Keys are the English text (or dotted ids for
 * templates); each locale folder merges the shell, grid and command string files of that locale.
 * This locale never changes the workbook's own language or number formats.
 */
import { de } from './locales/de';
import { en } from './locales/en';
import { es } from './locales/es';
import { fr } from './locales/fr';
import { zhCN } from './locales/zh-CN';

export type EditorLocale = 'en' | 'fr' | 'de' | 'es' | 'zh-CN';
/** A supported locale or any BCP 47 tag that maps onto one (`de-DE`, `zh-Hans`). */
export type EditorLocaleInput = EditorLocale | (string & {});
export const EDITOR_LOCALES: readonly EditorLocale[] = ['en', 'fr', 'de', 'es', 'zh-CN'];

/** Every display locale, keyed by its canonical code. `en` defines the key set. */
export const STRINGS: Readonly<Record<EditorLocale, Readonly<Record<string, string>>>> = {
	en,
	fr,
	de,
	es,
	'zh-CN': zhCN,
};

/**
 * Maps any BCP 47 tag to a supported display locale, falling back to English. Region variants
 * resolve to their language (`fr-CA`, `de-DE`, `es-MX`); `zh`, `zh-CN`, `zh-Hans` and `zh-SG` use
 * Simplified Chinese, while Traditional tags (`zh-TW`, `zh-HK`, `zh-Hant`) fall back to English.
 */
export function normalizeEditorLocale(value: string | null | undefined): EditorLocale {
	const parts = (value ?? '').trim().toLowerCase().split(/[-_]/);
	switch (parts[0]) {
		case 'fr':
		case 'de':
		case 'es':
			return parts[0];
		case 'zh':
			return parts.some((part) => ['hant', 'tw', 'hk', 'mo'].includes(part)) ? 'en' : 'zh-CN';
		default:
			return 'en';
	}
}

/** Fills `{name}` placeholders; unknown names are left as written. */
export function fillTemplate(text: string, vars?: Record<string, string | number>): string {
	if (!vars) return text;
	return text.replace(/\{(\w+)\}/g, (match, name: string) =>
		name in vars ? String(vars[name]) : match,
	);
}

/** Translates an English key (falling back to English, then to the key itself) and fills it. */
export function translate(
	locale: EditorLocale,
	key: string,
	vars?: Record<string, string | number>,
): string {
	const text = STRINGS[locale][key] ?? STRINGS.en[key] ?? key;
	return fillTemplate(text, vars);
}

/** A bound translator for one locale, the shape of `EditorContext.t`. */
export function translator(
	locale: () => EditorLocale,
): (key: string, vars?: Record<string, string | number>) => string {
	return (key, vars) => translate(locale(), key, vars);
}

/** Formats a number in the display locale (status bar statistics, zoom). */
export function formatNumber(locale: EditorLocale, value: number, digits = 10): string {
	try {
		return new Intl.NumberFormat(locale, { maximumFractionDigits: digits }).format(value);
	} catch {
		return String(value);
	}
}

/** Base of new sheet names in the UI language (`Sheet`, `Tabelle`, `Feuil`), as Excel uses. */
export function sheetBaseName(t: (key: string) => string): string {
	const key = 'sheet.default-name';
	const text = t(key);
	return text && text !== key ? text : 'Sheet';
}

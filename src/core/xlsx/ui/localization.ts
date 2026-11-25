/**
 * Display-language strings for the editor UI. Keys are the English text (or dotted ids for
 * templates); each locale folder merges the shell, grid and command string files of that locale.
 * This locale never changes the workbook's own language or number formats.
 */
import { createTranslator, interpolate, normalizeLocale } from '../../i18n/index';
import { en } from './locales/en/index';

export type EditorLocale = 'en' | 'fr' | 'de' | 'es' | 'zh-CN';
/** A supported locale or any BCP 47 tag that maps onto one (`de-DE`, `zh-Hans`). */
export type EditorLocaleInput = EditorLocale | (string & {});
export const EDITOR_LOCALES: readonly EditorLocale[] = ['en', 'fr', 'de', 'es', 'zh-CN'];

type EditorStrings = Readonly<Record<string, string>>;
type TranslatedLocale = Exclude<EditorLocale, 'en'>;

const loaders: Readonly<Record<TranslatedLocale, () => Promise<EditorStrings>>> = {
	fr: () => import('./locales/fr/index').then(({ fr }) => fr),
	de: () => import('./locales/de/index').then(({ de }) => de),
	es: () => import('./locales/es/index').then(({ es }) => es),
	'zh-CN': () => import('./locales/zh-CN/index').then(({ zhCN }) => zhCN),
};
const loaded: Partial<Record<TranslatedLocale, EditorStrings>> = {};
const pending = new Map<TranslatedLocale, Promise<EditorStrings>>();

/** The current dictionaries, with English available before another language loads. */
export const STRINGS: Readonly<Record<EditorLocale, EditorStrings>> = {
	en,
	get fr() {
		return loaded.fr ?? en;
	},
	get de() {
		return loaded.de ?? en;
	},
	get es() {
		return loaded.es ?? en;
	},
	get 'zh-CN'() {
		return loaded['zh-CN'] ?? en;
	},
};

/** Loads just the selected language. Callers can keep using English if loading fails. */
export function loadEditorStrings(locale: EditorLocale): Promise<EditorStrings> {
	if (locale === 'en') return Promise.resolve(en);
	const cached = loaded[locale];
	if (cached) return Promise.resolve(cached);
	const loading = pending.get(locale);
	if (loading) return loading;
	const request = loaders[locale]()
		.then((strings) => {
			loaded[locale] = strings;
			return strings;
		})
		.finally(() => pending.delete(locale));
	pending.set(locale, request);
	return request;
}

/**
 * Maps any BCP 47 tag to a supported display locale, falling back to English. Region variants
 * resolve to their language (`fr-CA`, `de-DE`, `es-MX`); `zh`, `zh-CN`, `zh-Hans` and `zh-SG` use
 * Simplified Chinese, while Traditional tags (`zh-TW`, `zh-HK`, `zh-Hant`) fall back to English.
 */
export function normalizeEditorLocale(value: string | null | undefined): EditorLocale {
	return normalizeLocale(value, EDITOR_LOCALES, 'en');
}

/** Fills `{name}` placeholders; unknown names are left as written. */
export function fillTemplate(text: string, vars?: Record<string, string | number>): string {
	return interpolate(text, vars);
}

/** Translates an English key (falling back to English, then to the key itself) and fills it. */
export function translate(
	locale: EditorLocale,
	key: string,
	vars?: Record<string, string | number>,
): string {
	return createTranslator(STRINGS[locale], [STRINGS.en])(key, vars);
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

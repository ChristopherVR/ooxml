// The Teams translation layer: one dictionary per locale (`locales/<locale>/*.ts`, English is the
// contract), the shared `ooxml-core/i18n` translator for lookup and `{placeholder}` filling, and
// `Intl` for plurals, dates and times. A missing key falls back to English, then to the key.
import {
	createTranslator,
	normalizeLocale,
	SUPPORTED_OFFICE_LOCALES,
	type MessageDictionary,
	type MessageParams,
	type OfficeLocale,
} from 'ooxml-core/i18n';
import { de } from './locales/de';
import { en } from './locales/en';
import { es } from './locales/es';
import { fr } from './locales/fr';
import { zhCN } from './locales/zh-CN';

export type TeamsLocale = OfficeLocale;
export const TEAMS_LOCALES = SUPPORTED_OFFICE_LOCALES;

/** Every key of the English dictionary, the contract of all locales. */
export type TeamsKey = keyof typeof en;
/** A plural message is stored as `<name>.one` and `<name>.other`; this is the shared `<name>`. */
export type TeamsPluralKey = TeamsKey extends infer K
	? K extends `${infer Name}.other`
		? Name
		: never
	: never;

export type TeamsTranslate = (key: TeamsKey, params?: MessageParams) => string;
export type TeamsPlural = (key: TeamsPluralKey, count: number, params?: MessageParams) => string;

export interface TeamsI18n {
	/** The resolved display locale (`en` when none was asked for). */
	readonly locale: TeamsLocale;
	/**
	 * The locale `Intl` formats dates and numbers in. Undefined while no locale was asked for, so
	 * a host that never sets one keeps the browser's own date and time conventions.
	 */
	readonly intl: string | undefined;
	readonly t: TeamsTranslate;
	/** `tn('chat.replies', 3)` picks `chat.replies.one` or `.other` by the locale's plural rules. */
	readonly tn: TeamsPlural;
}

/** Every locale's dictionary. */
export const TEAMS_MESSAGES: Readonly<Record<TeamsLocale, MessageDictionary>> = {
	en,
	de,
	es,
	fr,
	'zh-CN': zhCN,
};

/** Maps any BCP 47 tag (`de-AT`, `zh-Hans`, `fr_CA`) to a shipped locale, else English. */
export function normalizeTeamsLocale(input: string | null | undefined): TeamsLocale {
	return normalizeLocale(input, TEAMS_LOCALES, 'en');
}

const cache = new Map<string, TeamsI18n>();

/** The translator for a locale tag. Cached; an empty tag means English with the browser's dates. */
export function teamsI18n(input?: string | null): TeamsI18n {
	const asked = (input ?? '').trim();
	const hit = cache.get(asked);
	if (hit) return hit;
	const locale = normalizeTeamsLocale(asked);
	const translate = createTranslator(TEAMS_MESSAGES[locale], [TEAMS_MESSAGES.en]);
	const rules = new Intl.PluralRules(locale);
	const made: TeamsI18n = {
		locale,
		intl: asked ? locale : undefined,
		t: translate as TeamsTranslate,
		tn: (key, count, params) =>
			translate(`${key}.${rules.select(count) === 'one' ? 'one' : 'other'}`, {
				count,
				...params,
			}),
	};
	cache.set(asked, made);
	return made;
}

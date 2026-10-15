/**
 * The Visio ribbon's translator. The ribbon is built from English text (`ribbon-*.ts`), and its
 * dictionaries (`locales/`) are keyed by that English text, so a locale change translates the
 * built elements in place and a missing entry keeps the English wording. Dynamic wording is
 * handled by `{placeholder}` templates, the "label (shortcut)" and "label: reason" tooltip forms
 * and sentence-by-sentence lookup.
 */
import { interpolate, normalizeLocale } from 'ooxml-core/i18n';
import { deCommands } from './locales/de-commands';
import { deText } from './locales/de-text';
import { enCommands } from './locales/en-commands';
import { enText } from './locales/en-text';
import { esCommands } from './locales/es-commands';
import { esText } from './locales/es-text';
import { frCommands } from './locales/fr-commands';
import { frText } from './locales/fr-text';
import { zhCNCommands } from './locales/zh-CN-commands';
import { zhCNText } from './locales/zh-CN-text';

export const VISIO_LOCALES = ['en', 'de', 'es', 'fr', 'zh-CN'] as const;
export type VisioLocale = (typeof VISIO_LOCALES)[number];

/** Every dictionary, English first: its keys are the contract the other locales must meet. */
export const VISIO_RIBBON_STRINGS: Readonly<Record<VisioLocale, Readonly<Record<string, string>>>> =
	{
		en: { ...enCommands, ...enText },
		de: { ...deCommands, ...deText },
		es: { ...esCommands, ...esText },
		fr: { ...frCommands, ...frText },
		'zh-CN': { ...zhCNCommands, ...zhCNText },
	};

/** Maps any BCP 47 tag to a supported locale (`de-AT` is German, `zh-Hans` Chinese), else English. */
export function normalizeVisioLocale(value: string | null | undefined): VisioLocale {
	return normalizeLocale(value, VISIO_LOCALES, 'en');
}

const templateCache = new Map<VisioLocale, [RegExp, string, string[]][]>();

/** The `{name}` templates of a locale as [matcher, translation, placeholder names]. */
function templatesOf(locale: VisioLocale): [RegExp, string, string[]][] {
	let found = templateCache.get(locale);
	if (found) return found;
	found = [];
	for (const key of Object.keys(VISIO_RIBBON_STRINGS.en)) {
		if (!key.includes('{')) continue;
		const names: string[] = [];
		const source = key
			.split(/\{(\w+)\}/)
			.map((part, i) => {
				if (i % 2 === 0) return part.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
				names.push(part);
				return '(.+?)';
			})
			.join('');
		found.push([new RegExp(`^${source}$`), VISIO_RIBBON_STRINGS[locale][key] ?? key, names]);
	}
	templateCache.set(locale, found);
	return found;
}

/** Translates `text` exactly as written, or returns undefined. */
function exact(locale: VisioLocale, text: string): string | undefined {
	const dictionary = VISIO_RIBBON_STRINGS[locale];
	if (Object.hasOwn(dictionary, text)) return dictionary[text];
	for (const [pattern, translation, names] of templatesOf(locale)) {
		const match = pattern.exec(text);
		if (!match) continue;
		const values: Record<string, string> = {};
		names.forEach((name, i) => {
			const value = match[i + 1] ?? '';
			values[name] = translateRibbonText(locale, value);
		});
		return interpolate(translation, values);
	}
	return undefined;
}

/**
 * The English ribbon wording `text` in `locale`. Tooltips are `label (shortcut)`, `label: reason`
 * and `label: not available yet. reason`; anything unknown stays English.
 */
export function translateRibbonText(locale: VisioLocale, text: string): string {
	if (locale === 'en') return text;
	const direct = exact(locale, text);
	if (direct !== undefined) return direct;
	const unavailable = /^(.+?): not available yet\. (.*)$/s.exec(text);
	if (unavailable) {
		const template = VISIO_RIBBON_STRINGS[locale]['{label}: not available yet. {reason}'];
		return interpolate(template ?? text, {
			label: translateRibbonText(locale, unavailable[1] ?? ''),
			reason: translateRibbonText(locale, unavailable[2] ?? ''),
		});
	}
	const shortcut = /^(.+?) (\([^()]*\))$/.exec(text);
	if (shortcut) {
		const label = exact(locale, shortcut[1] ?? '');
		if (label !== undefined) return `${label} ${shortcut[2]}`;
	}
	const reason = /^(.+?): (.+)$/s.exec(text);
	if (reason) {
		const label = exact(locale, reason[1] ?? '');
		const rest = translateRibbonText(locale, reason[2] ?? '');
		if (label !== undefined) return `${label}: ${rest}`;
	}
	const sentences = text.match(/[^.]+\.(?:\s+|$)/g);
	if (sentences && sentences.length > 1 && sentences.join('') === text) {
		const parts = sentences.map((sentence) => exact(locale, sentence.trim()));
		if (parts.every((part) => part !== undefined)) return parts.join(' ');
	}
	return text;
}

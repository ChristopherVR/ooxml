import { createTranslator as createMessageTranslator } from 'ooxml-core/i18n';
import { keyToLabel, translationsEn } from './translations-en';

/**
 * Minimal i18n for the vanilla binding.
 *
 * The other bindings delegate to their framework's i18n runtime (react-i18next,
 * vue-i18n, ngx-translate) fed by the shared `pptx.*` dictionary. The vanilla
 * binding has no framework, so this module provides the one missing piece: a
 * `t(key, params)` lookup with `{{param}}` interpolation over the same shared
 * English dictionary, plus per-locale overrides supplied by the host.
 *
 * Resolution order: host dictionary for the active locale, then the built-in
 * English dictionary, then a humanised fallback derived from the key (shared
 * `keyToLabel`), so a missing key never renders as a raw `pptx.*` string.
 */

/** Translate a dotted `pptx.*` key with optional `{{param}}` interpolation. */
export type Translator = (key: string, params?: Record<string, string | number>) => string;

/** Locale to flat `key: message` dictionary map supplied by the host. */
export type TranslationMessages = Record<string, Record<string, string>>;

/**
 * Build a {@link Translator} for a locale. `messages[locale]` (when provided)
 * wins over the built-in English dictionary; English is always the fallback.
 */
export function createTranslator(locale = 'en', messages?: TranslationMessages): Translator {
	return createMessageTranslator(
		messages?.[locale],
		[locale !== 'en' ? messages?.en : undefined, translationsEn],
		{ placeholder: 'double', missing: keyToLabel },
	);
}

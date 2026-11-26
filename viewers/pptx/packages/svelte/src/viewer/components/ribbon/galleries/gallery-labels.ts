import type { Translator } from '../../../../i18n/translator';

/** A translator in the shape shared's `galleryItemLabel` expects. */
export type StrictTranslator = (
	key: string,
	params?: Readonly<Record<string, string | number>>,
) => string;

/**
 * Shared gallery helpers expect a missing entry to return its key so they can
 * use the fallback from the descriptor.
 */
export function strictTranslator(t: Translator): StrictTranslator {
	return (key, params) => t(key, params ? { ...params } : undefined, key);
}

/** `key` translated, or `fallback` when the dictionary has no entry for it. */
export function translatedOr(t: Translator, key: string | undefined, fallback: string): string {
	if (!key) {
		return fallback;
	}
	const out = strictTranslator(t)(key);
	return out === key ? fallback : out;
}

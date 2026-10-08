/** The five display languages every Office product UI ships. */
export const SUPPORTED_OFFICE_LOCALES = ['en', 'fr', 'de', 'es', 'zh-CN'] as const;

export type OfficeLocale = (typeof SUPPORTED_OFFICE_LOCALES)[number];

/** Subtags that select Traditional Chinese, which no product ships. */
const DEFAULT_REJECTED_SUBTAGS: readonly string[] = ['hant', 'tw', 'hk', 'mo'];

export interface NormalizeLocaleOptions {
	/** Subtags that make a tag fall back instead of matching a regional entry (`zh-TW`). */
	readonly rejectedSubtags?: readonly string[];
}

/**
 * Maps any BCP 47 tag (`_` or `-`, any case) to one of `supported`, else `fallback`. Resolution:
 * the exact tag; then its language alone (`de-AT` -> `de`); then a regional entry of the same
 * language (`zh`, `zh-Hans-CN`, `zh-SG` -> `zh-CN`) unless the tag carries a rejected subtag
 * (`zh-TW`, `zh-Hant`).
 */
export function normalizeLocale<T extends string>(
	input: string | null | undefined,
	supported: readonly T[],
	fallback: T,
	options: NormalizeLocaleOptions = {},
): T {
	const parts = (input ?? '').trim().toLowerCase().split(/[-_]/).filter(Boolean);
	const language = parts[0];
	if (!language) return fallback;
	const find = (tag: string): T | undefined => supported.find((item) => item.toLowerCase() === tag);
	const exact = find(parts.join('-'));
	if (exact) return exact;
	const base = find(language);
	if (base) return base;
	const rejected = options.rejectedSubtags ?? DEFAULT_REJECTED_SUBTAGS;
	if (parts.slice(1).some((part) => rejected.includes(part))) return fallback;
	return (
		supported.find((item) => item.toLowerCase().split('-')[0] === language && item.includes('-')) ??
		fallback
	);
}

/** `normalizeLocale` over the five shared Office locales, falling back to English. */
export function normalizeOfficeLocale(input: string | null | undefined): OfficeLocale {
	return normalizeLocale(input, SUPPORTED_OFFICE_LOCALES, 'en');
}

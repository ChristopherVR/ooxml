import type { MessageDictionary } from './dictionary';

export type MessageParams = Readonly<Record<string, string | number>>;

/** Translates a key, filling placeholders from `params`. */
export type Translate = (key: string, params?: MessageParams) => string;

export interface TranslatorOptions {
	/** Placeholder syntax: `{name}` (default) or `{{name}}` (whitespace allowed inside). */
	readonly placeholder?: 'single' | 'double';
	/** Text for a key found in no dictionary; defaults to the key itself. */
	readonly missing?: (key: string) => string;
}

const SINGLE = /\{(\w+)\}/g;
const DOUBLE = /\{\{\s*(\w+)\s*\}\}/g;

/** Fills placeholders from `params`; unknown names are left as written. */
export function interpolate(
	template: string,
	params?: MessageParams,
	placeholder: 'single' | 'double' = 'single',
): string {
	if (!params) return template;
	const pattern = placeholder === 'double' ? DOUBLE : SINGLE;
	return template.replace(pattern, (match, name: string) =>
		Object.prototype.hasOwnProperty.call(params, name) && params[name] !== undefined
			? String(params[name])
			: match,
	);
}

/**
 * Builds a translator. Resolution: `messages`, then each of `fallbacks` in order (English
 * first), then `options.missing(key)`, else the key itself.
 */
export function createTranslator(
	messages: MessageDictionary | undefined,
	fallbacks: readonly (MessageDictionary | undefined)[] = [],
	options: TranslatorOptions = {},
): Translate {
	const placeholder = options.placeholder ?? 'single';
	return (key, params) => {
		let text = messages?.[key];
		for (let i = 0; text === undefined && i < fallbacks.length; i++) text = fallbacks[i]?.[key];
		return interpolate(text ?? options.missing?.(key) ?? key, params, placeholder);
	};
}

/**
 * `translate(key, params)` when the binding's dictionary has `key`, else `fallback`. Every
 * binding's translator answers a miss with the key itself, which is what this checks for, so a
 * host with no dictionary entry still shows the English name.
 */
export function translateOr(
	translate: Translate | undefined,
	key: string,
	fallback: string,
	params?: MessageParams,
): string {
	if (!translate) return fallback;
	const text = translate(key, params);
	return text && text !== key ? text : fallback;
}

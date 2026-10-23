import { interpolate } from 'ooxml-core/i18n';
import type { MessageParams } from 'ooxml-core/i18n';
import { keyToLabel, translationsEn } from 'ooxml-ui/pptx/i18n';
import { useMemo } from 'react';
import type { ReactElement, ReactNode } from 'react';

import { PptxI18nContext } from './context';
import type { PptxTranslate, PptxTranslateParams, ViewerI18nState, ViewerT } from './types';

export interface PptxI18nProviderProps {
	/**
	 * Translates a viewer key; see {@link PptxTranslate}. Keep it referentially stable (a
	 * module-level function or `useCallback`): the viewer re-translates, re-rendering and re-running
	 * every translated effect, whenever `translate` or `locale` changes. To pick up messages loaded
	 * later for the same language, pass a new function.
	 */
	readonly translate: PptxTranslate;
	/**
	 * The host's active language: what the viewer reports and File > Options > Language shows as
	 * selected. A change re-translates the viewer, also when `translate` stays the same.
	 */
	readonly locale: string;
	readonly children?: ReactNode;
}

/**
 * `interpolate` fills a placeholder with `String(value)`, which is also how i18next renders a
 * value that is not a string or a number. `MessageParams` only declares strings and numbers, so
 * this cast widens nothing at runtime: other values are stringified, never dropped.
 */
const asMessageParams = (params: PptxTranslateParams | undefined): MessageParams | undefined =>
	params as MessageParams | undefined;

/**
 * The viewer's `t` over a host translator. A host miss (the key itself, an empty string, `null`,
 * `undefined`, or a throw) falls back to the built-in English text, then the call's
 * `defaultValue` (an options field or the string overload), then a label derived from the key, so
 * a raw `pptx.*` key never shows. `{{name}}` placeholders left in the text are filled from the
 * params, so a host can return the shipped dictionaries as they are.
 */
function withEnglishFallback(translate: PptxTranslate): ViewerT {
	return (key, defaultOrParams) => {
		const params: PptxTranslateParams | undefined =
			typeof defaultOrParams === 'string' ? { defaultValue: defaultOrParams } : defaultOrParams;
		let text: string | null | undefined;
		try {
			text = params === undefined ? translate(key) : translate(key, params);
		} catch {
			// react-i18next never throws; one failing key must not take the editor down.
			text = undefined;
		}
		const found = text && text !== key ? text : (translationsEn[key] ?? params?.defaultValue);
		return typeof found === 'string'
			? interpolate(found, asMessageParams(params), 'double')
			: keyToLabel(key);
	};
}

/**
 * Translates every viewer component below it (`PowerPointViewer` and the standalone building
 * blocks) through the host's own i18n library instead of react-i18next. A `PowerPointViewer`
 * reads the provider nearest above it; a provider rendered inside the viewer is ignored, so the
 * whole viewer shares one language. React only: the other bindings translate through their
 * framework's i18n library or a `messages` dictionary.
 */
export function PptxI18nProvider({
	translate,
	locale,
	children,
}: PptxI18nProviderProps): ReactElement {
	const state = useMemo<ViewerI18nState>(
		() => ({ t: withEnglishFallback(translate), locale }),
		[translate, locale],
	);
	return <PptxI18nContext.Provider value={state}>{children}</PptxI18nContext.Provider>;
}

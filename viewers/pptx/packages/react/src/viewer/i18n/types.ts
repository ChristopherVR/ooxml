/** What the viewer passes `translate`: exactly the params of the call. */
export interface PptxTranslateParams {
	readonly defaultValue?: string;
	readonly count?: number;
	readonly [name: string]: unknown;
}

/**
 * A host translator for {@link PptxI18nProvider}: the text for a viewer key (`pptx.*`) in the
 * host's active language. Return the key itself, an empty string, `null` or `undefined` for "no
 * translation": the viewer then shows its built-in English text.
 */
export type PptxTranslate = (
	key: string,
	params?: PptxTranslateParams,
) => string | null | undefined;

/** Values a translation may interpolate (`{{name}}` in the shipped dictionaries). */
export type ViewerTranslateParams = PptxTranslateParams;

/**
 * The translate function viewer components call. It keeps both call shapes the viewer uses:
 * `t(key, { defaultValue, ...params })` and the string overload `t(key, 'Default text')`.
 */
export type ViewerT = (key: string, options?: ViewerTranslateParams | string) => string;

/** The resolved translation state every viewer component reads. */
export interface ViewerI18nState {
	readonly t: ViewerT;
	/** Active language of the translations `t` returns. */
	readonly locale: string;
}

/** Values a translation may interpolate (`{{name}}` in the shipped dictionaries). */
export type ViewerTranslateParams = Readonly<Record<string, unknown>>;

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

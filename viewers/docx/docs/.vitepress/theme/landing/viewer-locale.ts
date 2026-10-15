/**
 * Maps a VitePress site language (`useData().lang`, e.g. `'fr-FR'`) to the
 * locale id the embedded Word demo apps accept through `?locale=`. A small
 * self-contained lookup, so the docs site needs no workspace package at build
 * time. Keep the keys in sync with the `lang` values in `locales/*.ts`.
 */
export type ViewerLocaleId = 'en' | 'fr' | 'es' | 'de' | 'zh-CN';

const VIEWER_LOCALE_BY_LANG: Record<string, ViewerLocaleId> = {
	'en-US': 'en',
	'fr-FR': 'fr',
	'es-ES': 'es',
	'de-DE': 'de',
	'zh-CN': 'zh-CN',
};

/** Resolves the embedded viewer's locale for the given VitePress site `lang`. */
export function resolveViewerLocale(lang: string): ViewerLocaleId {
	return VIEWER_LOCALE_BY_LANG[lang] ?? 'en';
}

import { normalizeLocale } from 'ooxml-core/i18n';
import { readStoredViewerPrefs, writeStoredViewerPrefs } from 'ooxml-ui/pptx';
import { LOCALE_CATALOG } from 'ooxml-ui/pptx/i18n';
import type { LocaleCatalogEntry } from 'ooxml-ui/pptx/i18n';
import { useCallback, useContext, useEffect, useMemo, useState } from 'react';
import type { ReactElement, ReactNode } from 'react';

import { PptxI18nContext, ViewerI18nContext, ViewerLocaleControlsContext } from './context';
import type { ViewerLocaleControls } from './context';
import { useHostI18next } from './react-i18next-source';
import type { ViewerI18nState } from './types';

/** The `PowerPointViewer` props that decide its language. */
export interface ViewerI18nProps {
	readonly defaultLocale?: string | undefined;
	readonly availableLocales?: LocaleCatalogEntry[] | undefined;
	readonly onLocaleChange?: ((code: string) => void) | undefined;
}

const NO_LOCALES: LocaleCatalogEntry[] = [];

const toCatalogEntry = (code: string): LocaleCatalogEntry =>
	LOCALE_CATALOG.find((entry) => entry.code === code) ?? { code, label: code, nativeLabel: code };

/** What a viewer root provides: the translation state and the Language tab controls. */
interface ViewerLanguage {
	readonly state: ViewerI18nState;
	readonly controls: ViewerLocaleControls;
}

/**
 * The language from the host's react-i18next instance: the viewer never bundles an i18n
 * instance, it only calls `changeLanguage` on the host's.
 */
function useHostLanguage({
	defaultLocale,
	availableLocales,
	onLocaleChange,
}: ViewerI18nProps): ViewerLanguage {
	const { t, i18n } = useHostI18next();
	const [localeCode, setLocaleCode] = useState<string>(
		() => defaultLocale ?? readStoredViewerPrefs().localeCode ?? 'en',
	);
	// A persisted non-English choice needs to actually take effect on reload; the viewer only
	// applies it itself when the host has not taken over locale handling via `onLocaleChange`.
	useEffect(() => {
		if (localeCode !== 'en' && !onLocaleChange) {
			void i18n.changeLanguage(localeCode);
		}
		// Intentionally run once on mount: later changes flow through selectLocale below.
		// oxlint-disable-next-line react-hooks/exhaustive-deps
	}, []);
	const locales = useMemo<LocaleCatalogEntry[]>(() => {
		if (availableLocales) {
			return availableLocales;
		}
		// Introspect what the host's i18n instance actually has dictionaries for, rather than
		// assuming every LOCALE_CATALOG entry is wired up.
		const registeredCodes = i18n.options.resources
			? Object.keys(i18n.options.resources)
			: (i18n.languages ?? ['en']);
		return registeredCodes.map(toCatalogEntry);
	}, [availableLocales, i18n]);
	const selectLocale = useCallback(
		(code: string) => {
			setLocaleCode(code);
			if (onLocaleChange) {
				onLocaleChange(code);
			} else {
				void i18n.changeLanguage(code);
				writeStoredViewerPrefs({ localeCode: code });
			}
		},
		[onLocaleChange, i18n],
	);

	const language = i18n.resolvedLanguage ?? i18n.language ?? localeCode;
	const state = useMemo<ViewerI18nState>(() => ({ t, locale: language }), [t, language]);
	const controls = useMemo<ViewerLocaleControls>(
		() => ({ localeCode, locales, selectLocale, languageSelectable: true }),
		[localeCode, locales, selectLocale],
	);
	return { state, controls };
}

/**
 * The language from the nearest `PptxI18nProvider`: its `translate` and `locale`. The host owns
 * the language, so the viewer neither stores a choice nor changes it: the Language tab lists
 * `availableLocales` and a pick calls `onLocaleChange`. Without `onLocaleChange` and two or more
 * `availableLocales` the tab is left out.
 */
function useProvidedLanguage({
	availableLocales,
	onLocaleChange,
}: ViewerI18nProps): ViewerLanguage {
	// Only chosen below when a provider is present.
	const state = useContext(PptxI18nContext) as ViewerI18nState;
	const { locale } = state;
	const locales = availableLocales ?? NO_LOCALES;
	const controls = useMemo<ViewerLocaleControls>(
		() => ({
			// `fr-CA` selects an offered `fr`: the host's locale need not be one of the codes.
			localeCode: normalizeLocale(
				locale,
				locales.map((entry) => entry.code),
				locale,
			),
			locales,
			selectLocale: (code) => onLocaleChange?.(code),
			languageSelectable: onLocaleChange !== undefined && locales.length > 1,
		}),
		[locale, locales, onLocaleChange],
	);
	return { state, controls };
}

/** The context providers of one language. */
function provideLanguage({ state, controls }: ViewerLanguage, children: ReactNode): ReactElement {
	return (
		<ViewerI18nContext.Provider value={state}>
			<ViewerLocaleControlsContext.Provider value={controls}>
				{children}
			</ViewerLocaleControlsContext.Provider>
		</ViewerI18nContext.Provider>
	);
}

/** A viewer root's props: the language props of `PowerPointViewer`, and the viewer body. */
export interface ViewerI18nRootProps extends ViewerI18nProps {
	readonly children: ReactNode;
}

function HostLanguageRoot({ children, ...props }: ViewerI18nRootProps): ReactElement {
	return provideLanguage(useHostLanguage(props), children);
}

function ProvidedLanguageRoot({ children, ...props }: ViewerI18nRootProps): ReactElement {
	return provideLanguage(useProvidedLanguage(props), children);
}

/**
 * Resolves the viewer's language once, above every hook that translates, and provides it to
 * the components below: from the nearest `PptxI18nProvider` when there is one, else from the
 * host's react-i18next instance. Adding or removing a provider remounts the viewer.
 */
export function ViewerI18nRoot(props: ViewerI18nRootProps): ReactElement {
	const Root = useContext(PptxI18nContext) ? ProvidedLanguageRoot : HostLanguageRoot;
	return <Root {...props} />;
}

/** File > Options > Language state of the nearest viewer root. */
export function useViewerLocaleControls(): ViewerLocaleControls {
	return useContext(ViewerLocaleControlsContext);
}

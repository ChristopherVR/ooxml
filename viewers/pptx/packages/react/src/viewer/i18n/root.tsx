import { readStoredViewerPrefs, writeStoredViewerPrefs } from 'ooxml-ui/pptx';
import { LOCALE_CATALOG } from 'ooxml-ui/pptx/i18n';
import type { LocaleCatalogEntry } from 'ooxml-ui/pptx/i18n';
import { useCallback, useContext, useEffect, useMemo, useState } from 'react';
import type { ReactElement, ReactNode } from 'react';

import { ViewerI18nContext, ViewerLocaleControlsContext } from './context';
import type { ViewerLocaleControls } from './context';
import { useHostI18next } from './react-i18next-source';
import type { ViewerI18nState } from './types';

/** The `PowerPointViewer` props that decide its language. */
export interface ViewerI18nProps {
	readonly defaultLocale?: string | undefined;
	readonly availableLocales?: LocaleCatalogEntry[] | undefined;
	readonly onLocaleChange?: ((code: string) => void) | undefined;
}

/** A viewer root's props: the language props of `PowerPointViewer`, and the viewer body. */
export interface ViewerI18nRootProps extends ViewerI18nProps {
	readonly children: ReactNode;
}

const toCatalogEntry = (code: string): LocaleCatalogEntry =>
	LOCALE_CATALOG.find((entry) => entry.code === code) ?? { code, label: code, nativeLabel: code };

/**
 * Resolves the viewer's language once, above every hook that translates, and provides it to
 * the components below. The source is the host's react-i18next instance: the viewer never
 * bundles an i18n instance, it only calls `changeLanguage` on the host's.
 */
export function ViewerI18nRoot({
	defaultLocale,
	availableLocales,
	onLocaleChange,
	children,
}: ViewerI18nRootProps): ReactElement {
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
		() => ({ localeCode, locales, selectLocale }),
		[localeCode, locales, selectLocale],
	);
	return (
		<ViewerI18nContext.Provider value={state}>
			<ViewerLocaleControlsContext.Provider value={controls}>
				{children}
			</ViewerLocaleControlsContext.Provider>
		</ViewerI18nContext.Provider>
	);
}

/** File > Options > Language state of the nearest viewer root. */
export function useViewerLocaleControls(): ViewerLocaleControls {
	return useContext(ViewerLocaleControlsContext);
}

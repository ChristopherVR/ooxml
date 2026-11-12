import { normalizeLocale } from 'ooxml-core/i18n';
import { readStoredViewerPrefs, writeStoredViewerPrefs } from 'ooxml-ui/pptx';
import { LOCALE_CATALOG } from 'ooxml-ui/pptx/i18n';
import type { LocaleCatalogEntry } from 'ooxml-ui/pptx/i18n';
import { useCallback, useContext, useEffect, useMemo, useState } from 'react';
import type { ReactElement, ReactNode } from 'react';

import { PptxI18nContext, ViewerI18nContext, ViewerLocaleControlsContext } from './context';
import type { ViewerLocaleControls } from './context';
import {
	hostInstanceLanguage,
	registeredHostI18next,
	useRegisteredHostI18next,
} from './host-i18next';
import { BUILT_IN_ENGLISH } from './provider';
import type { ViewerI18nState } from './types';

/** The `PowerPointViewer` props that decide its language. */
export interface ViewerI18nProps {
	readonly defaultLocale?: string | undefined;
	readonly availableLocales?: LocaleCatalogEntry[] | undefined;
	readonly onLocaleChange?: ((code: string) => void) | undefined;
}

const NO_LOCALES: LocaleCatalogEntry[] = [];

/**
 * Neither a provider nor a configured react-i18next source: the built-in English text, and no
 * Language tab, since there is nothing to switch to. The viewer neither reads nor stores a choice.
 */
const BUILT_IN_CONTROLS: ViewerLocaleControls = {
	localeCode: 'en',
	locales: NO_LOCALES,
	selectLocale: () => {},
	languageSelectable: false,
};

const toCatalogEntry = (code: string): LocaleCatalogEntry =>
	LOCALE_CATALOG.find((entry) => entry.code === code) ?? { code, label: code, nativeLabel: code };

/** What a viewer root provides: the translation state and the Language tab controls. */
interface ViewerLanguage {
	readonly state: ViewerI18nState;
	readonly controls: ViewerLocaleControls;
}

/**
 * The language from the host's react-i18next instance: the viewer never bundles an i18n
 * instance, it only calls `changeLanguage` on the host's. While react-i18next has no configured
 * instance this is the built-in English language; every hook still runs, so the hook calls stay
 * the same if react-i18next gets an instance later.
 */
function useHostLanguage({
	defaultLocale,
	availableLocales,
	onLocaleChange,
}: ViewerI18nProps): ViewerLanguage {
	const host = useRegisteredHostI18next();
	const t = host?.t;
	const i18n = host?.i18n;
	const readsInstance = i18n !== undefined;
	const [localeCode, setLocaleCode] = useState<string>(
		() => defaultLocale ?? readStoredViewerPrefs().localeCode ?? 'en',
	);
	// A persisted non-English choice needs to actually take effect on reload; the viewer only
	// applies it itself when the host has not taken over locale handling via `onLocaleChange`.
	useEffect(() => {
		if (i18n && localeCode !== 'en' && !onLocaleChange) {
			void i18n.changeLanguage?.(localeCode);
		}
		// Intentionally run on mount, and once more if react-i18next only gets an instance after
		// the viewer mounted and its parent re-renders it (no event fires when an instance first
		// appears): later changes flow through selectLocale below.
		// oxlint-disable-next-line react-hooks/exhaustive-deps
	}, [readsInstance]);
	const locales = useMemo<LocaleCatalogEntry[]>(() => {
		if (availableLocales) {
			return availableLocales;
		}
		if (!i18n) {
			return NO_LOCALES;
		}
		// Introspect what the host's i18n instance actually has dictionaries for, rather than
		// assuming every LOCALE_CATALOG entry is wired up.
		const registeredCodes = i18n.options?.resources
			? Object.keys(i18n.options.resources)
			: Array.isArray(i18n.languages)
				? i18n.languages
				: ['en'];
		return registeredCodes.map(toCatalogEntry);
	}, [availableLocales, i18n]);
	const selectLocale = useCallback(
		(code: string) => {
			setLocaleCode(code);
			if (onLocaleChange) {
				onLocaleChange(code);
			} else if (i18n) {
				void i18n.changeLanguage?.(code);
				writeStoredViewerPrefs({ localeCode: code });
			}
		},
		[onLocaleChange, i18n],
	);

	const language = i18n ? (hostInstanceLanguage(i18n) ?? localeCode) : 'en';
	const state = useMemo<ViewerI18nState>(
		() => (t ? { t, locale: language } : BUILT_IN_ENGLISH),
		[t, language],
	);
	const controls = useMemo<ViewerLocaleControls>(
		() =>
			readsInstance
				? { localeCode, locales, selectLocale, languageSelectable: true }
				: BUILT_IN_CONTROLS,
		[readsInstance, localeCode, locales, selectLocale],
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

function BuiltInLanguageRoot({ children }: ViewerI18nRootProps): ReactElement {
	return provideLanguage({ state: BUILT_IN_ENGLISH, controls: BUILT_IN_CONTROLS }, children);
}

function ProvidedLanguageRoot({ children, ...props }: ViewerI18nRootProps): ReactElement {
	return provideLanguage(useProvidedLanguage(props), children);
}

/**
 * Resolves the viewer's language once, above every hook that translates, and provides it to
 * the components below: from the nearest `PptxI18nProvider` when there is one, else from the
 * host's react-i18next instance when an entry registered it (`./host-i18next`) and react-i18next
 * has a configured instance, else the built-in English text. Adding or removing a provider remounts the viewer.
 */
export function ViewerI18nRoot(props: ViewerI18nRootProps): ReactElement {
	// Read once per mounted root, so a source registered later never swaps its body.
	// oxlint-disable-next-line react/hook-use-state -- read once on mount and never set
	const [readsHost] = useState(() => registeredHostI18next() !== null);
	const FromHost = readsHost ? HostLanguageRoot : BuiltInLanguageRoot;
	const Root = useContext(PptxI18nContext) ? ProvidedLanguageRoot : FromHost;
	return <Root {...props} />;
}

/** File > Options > Language state of the nearest viewer root. */
export function useViewerLocaleControls(): ViewerLocaleControls {
	return useContext(ViewerLocaleControlsContext);
}

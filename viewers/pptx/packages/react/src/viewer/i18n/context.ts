import type { LocaleCatalogEntry } from 'ooxml-ui/pptx/i18n';
import { createContext } from 'react';

import type { ViewerI18nState } from './types';

/**
 * The translation state of the nearest viewer root. `null` outside one: a standalone building
 * block (`Toolbar`, `SlideCanvas`, ...) then translates through the host's react-i18next.
 */
export const ViewerI18nContext = /* @__PURE__ */ createContext<ViewerI18nState | null>(null);

/**
 * The translation state of the nearest `PptxI18nProvider`, `null` without one. Inside a provider
 * the viewer never reads react-i18next.
 */
export const PptxI18nContext = /* @__PURE__ */ createContext<ViewerI18nState | null>(null);

/** What File > Options > Language reads and calls. */
export interface ViewerLocaleControls {
	/** The code the Language tab shows as selected. */
	readonly localeCode: string;
	/** The choices the Language tab offers. */
	readonly locales: LocaleCatalogEntry[];
	/** Applies a pick from the Language tab. */
	readonly selectLocale: (code: string) => void;
	/**
	 * `false` leaves the Language category out of File > Options. Inside a `PptxI18nProvider` it is
	 * `true` only with `onLocaleChange` and two or more `availableLocales`; elsewhere always `true`.
	 */
	readonly languageSelectable: boolean;
}

const NO_LOCALE_CONTROLS: ViewerLocaleControls = {
	localeCode: 'en',
	locales: [],
	selectLocale: () => {},
	languageSelectable: true,
};

export const ViewerLocaleControlsContext =
	/* @__PURE__ */ createContext<ViewerLocaleControls>(NO_LOCALE_CONTROLS);

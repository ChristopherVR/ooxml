import type { LocaleCatalogEntry } from 'ooxml-ui/pptx/i18n';
import { createContext } from 'react';

import type { ViewerI18nState } from './types';

/**
 * The translation state of the nearest viewer root. `null` outside one: a standalone building
 * block (`Toolbar`, `SlideCanvas`, ...) then translates through the host's react-i18next.
 */
export const ViewerI18nContext = /* @__PURE__ */ createContext<ViewerI18nState | null>(null);

/** What File > Options > Language reads and calls. */
export interface ViewerLocaleControls {
	/** The code the Language tab shows as selected. */
	readonly localeCode: string;
	/** The choices the Language tab offers; it hides itself with one or none. */
	readonly locales: LocaleCatalogEntry[];
	/** Applies a pick from the Language tab. */
	readonly selectLocale: (code: string) => void;
}

const NO_LOCALE_CONTROLS: ViewerLocaleControls = {
	localeCode: 'en',
	locales: [],
	selectLocale: () => {},
};

export const ViewerLocaleControlsContext =
	/* @__PURE__ */ createContext<ViewerLocaleControls>(NO_LOCALE_CONTROLS);

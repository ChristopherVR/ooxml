import { keyToLabel } from 'ooxml-ui/pptx/i18n';
import { useContext, useMemo, useState } from 'react';

import { PptxI18nContext, ViewerI18nContext } from './context';
import { useHostI18next } from './react-i18next-source';
import type { ViewerI18nState } from './types';

/** Only reachable if a context a component mounted under went away, which React never does. */
const LABELS_ONLY: ViewerI18nState = { t: (key) => keyToLabel(key), locale: 'en' };

/** The host's react-i18next translation, as a viewer translation state. */
function useHostTranslation(): ViewerI18nState {
	const { t, i18n } = useHostI18next();
	const language = i18n.resolvedLanguage ?? i18n.language ?? 'en';
	return useMemo<ViewerI18nState>(() => ({ t, locale: language }), [t, language]);
}

/**
 * The viewer's translation hook. Same call shape as the react-i18next hook it replaces
 * (`const { t } = useTranslation()`), so a call site only changes its import.
 *
 * Inside a viewer root the state comes from that root, so a `PptxI18nProvider` rendered inside
 * the viewer is ignored. Outside one (a standalone `Toolbar`, `SlideCanvas`, ...) it is the
 * nearest `PptxI18nProvider`, else the host's react-i18next instance. Only a component with
 * neither calls react-i18next, so one viewer subscribes to i18next once, in its root.
 *
 * A leaf module on purpose: specs mock exactly this file, never a barrel the root also uses.
 */
export function useTranslation(): ViewerI18nState {
	const fromRoot = useContext(ViewerI18nContext);
	const provided = useContext(PptxI18nContext);
	const resolved = fromRoot ?? provided;
	// Decided once per mounted component, so its hook calls never change. Both contexts are
	// provided unconditionally, so a mounted component never gains or loses one anyway.
	// oxlint-disable-next-line react/hook-use-state -- read once on mount and never set
	const [readsHost] = useState(resolved === null);
	// oxlint-disable-next-line react-hooks/rules-of-hooks -- `readsHost` is fixed for the component's lifetime
	const host = readsHost ? useHostTranslation() : null;
	return resolved ?? host ?? LABELS_ONLY;
}

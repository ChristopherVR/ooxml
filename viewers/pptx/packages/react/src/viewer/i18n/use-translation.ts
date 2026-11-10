import { useContext, useMemo, useState } from 'react';

import { PptxI18nContext, ViewerI18nContext } from './context';
import { BUILT_IN_ENGLISH } from './provider';
import { hostInstanceLanguage, useHostI18next } from './react-i18next-source';
import type { ViewerI18nState } from './types';

/**
 * The host's react-i18next translation, as a viewer translation state; `null` while react-i18next
 * has no configured instance. The memo runs either way, so the hook calls never change.
 */
function useHostTranslation(): ViewerI18nState | null {
	const host = useHostI18next();
	const t = host?.t;
	const language = (host && hostInstanceLanguage(host.i18n)) ?? 'en';
	return useMemo<ViewerI18nState | null>(() => (t ? { t, locale: language } : null), [t, language]);
}

/**
 * The viewer's translation hook. Same call shape as the react-i18next hook it replaces
 * (`const { t } = useTranslation()`), so a call site only changes its import.
 *
 * Inside a viewer root the state comes from that root, so a `PptxI18nProvider` rendered inside
 * the viewer is ignored. Outside one (a standalone `Toolbar`, `SlideCanvas`, ...) it is the
 * nearest `PptxI18nProvider`, else the host's react-i18next instance when react-i18next has one,
 * else the built-in English text. Only a component with neither context calls react-i18next, so
 * one viewer subscribes to i18next once, in its root.
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
	return resolved ?? host ?? BUILT_IN_ENGLISH;
}

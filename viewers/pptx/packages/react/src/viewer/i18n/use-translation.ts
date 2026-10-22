import { useContext, useMemo } from 'react';

import { ViewerI18nContext } from './context';
import { useHostI18next } from './react-i18next-source';
import type { ViewerI18nState } from './types';

/**
 * The viewer's translation hook. Same call shape as the react-i18next hook it replaces
 * (`const { t } = useTranslation()`), so a call site only changes its import.
 *
 * Inside a viewer root the state comes from that root. Outside one (a standalone `Toolbar`,
 * `SlideCanvas`, ...) it is the host's react-i18next instance. The react-i18next hook runs
 * either way, so the hook order never depends on the tree.
 *
 * A leaf module on purpose: specs mock exactly this file, never a barrel the root also uses.
 */
export function useTranslation(): ViewerI18nState {
	const { t, i18n } = useHostI18next();
	const language = i18n.resolvedLanguage ?? i18n.language ?? 'en';
	const host = useMemo<ViewerI18nState>(() => ({ t, locale: language }), [t, language]);
	return useContext(ViewerI18nContext) ?? host;
}

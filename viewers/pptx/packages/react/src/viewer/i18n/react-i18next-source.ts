/**
 * The host's react-i18next instance as a translation source. This is the only module of the
 * package that imports `react-i18next`; every component translates through
 * `./use-translation` instead.
 */
import { useTranslation as useReactI18nextTranslation } from 'react-i18next';

import type { ViewerT } from './types';

/** The parts of an i18next instance the viewer reads (the File > Options > Language tab). */
export interface HostI18nextInstance {
	readonly language?: string;
	readonly resolvedLanguage?: string;
	readonly languages?: readonly string[];
	readonly options: { readonly resources?: Readonly<Record<string, unknown>> };
	changeLanguage(code: string): unknown;
}

export interface HostI18next {
	readonly t: ViewerT;
	readonly i18n: HostI18nextInstance;
}

/**
 * react-i18next's own `useTranslation`, called unconditionally, so Suspense, `ready` and the
 * `languageChanged` re-render behave exactly as react-i18next defines them.
 */
export function useHostI18next(): HostI18next {
	const { t, i18n } = useReactI18nextTranslation();
	// i18next's `TFunction` accepts every call shape the viewer uses (an options object with
	// `defaultValue` and `{{name}}` params, or a default string); only its overloads differ.
	return { t: t as unknown as ViewerT, i18n: i18n as unknown as HostI18nextInstance };
}

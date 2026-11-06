/**
 * Where the viewer finds the host's react-i18next instance when no `PptxI18nProvider` is above
 * it. This module never imports react-i18next: `./react-i18next-source` registers itself here when
 * it is evaluated, and only the entries that fall back to react-i18next import it
 * (`pptx-react-viewer`, `pptx-react-viewer/viewer` and `pptx-react-viewer/internals`). An app
 * that imports none of them, such as one using `pptx-react-viewer/no-i18next`, leaves the slot
 * empty, and the viewer then shows its built-in English text outside a provider. So does a
 * registered source while react-i18next has no configured instance.
 */
import type { HostI18next, HostI18nextHook, HostI18nextInstance } from './types';

let registered: HostI18nextHook | null = null;

/**
 * Makes `useHostI18next` the fallback translation source. The first registration wins, so
 * evaluating the registering module again (another entry, a second bundle chunk) changes nothing,
 * and once registered the source never changes.
 */
export function registerHostI18next(useHostI18next: HostI18nextHook): void {
	registered ??= useHostI18next;
}

/** The registered fallback source, `null` when no entry that registers one was imported. */
export function registeredHostI18next(): HostI18nextHook | null {
	return registered;
}

/**
 * The active language of a host instance: `resolvedLanguage`, else `language`. `undefined` when
 * neither is a string, as with a test double or a Proxy auto-mock.
 */
export function hostInstanceLanguage(i18n: HostI18nextInstance): string | undefined {
	const { resolvedLanguage, language } = i18n;
	if (typeof resolvedLanguage === 'string') {
		return resolvedLanguage;
	}
	return typeof language === 'string' ? language : undefined;
}

/**
 * Calls the registered source. Only for a component that found one registered when it mounted:
 * a registration is never replaced, so it calls the same hook on every render. `null` when
 * react-i18next has no configured instance; the caller then uses the built-in English text.
 */
export function useRegisteredHostI18next(): HostI18next | null {
	if (!registered) {
		throw new Error('No react-i18next source is registered');
	}
	return registered();
}

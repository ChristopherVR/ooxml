/**
 * The host's react-i18next instance as a translation source. This is the only module of the
 * package that imports `react-i18next`; every component translates through
 * `./use-translation` instead. Only a viewer root and a standalone building block outside any
 * `PptxI18nProvider` call it.
 */
import { useTranslation as useReactI18nextTranslation } from 'react-i18next';

import type { HostI18next, HostI18nextInstance, ViewerT } from './types';

/** An object with no own enumerable keys whose prototype is `Object.prototype` (any realm) or `null`. */
function isEmptyPlainObject(value: unknown): boolean {
	if (typeof value !== 'object' || value === null) {
		return false;
	}
	const prototype: unknown = Object.getPrototypeOf(value);
	// `Object.prototype` of any realm: its own prototype is `null`.
	const plain = prototype === null || Object.getPrototypeOf(prototype as object) === null;
	return plain && Object.keys(value).length === 0;
}

/**
 * Whether `useTranslation` handed back react-i18next's stand-in for a missing instance: `ready`
 * is `false` and `i18n` is an empty plain object (react-i18next 15 and 16.0 return `{}`; later
 * versions a copy of it). Any instance it resolved, i18next's or a hand-built one, has own fields
 * (react-i18next writes `reportNamespaces` onto it) or a class prototype. Anything else is a
 * host's mock of `useTranslation`, such as `() => ({ t })`, `{ t, i18n: {} }` or a Proxy
 * auto-mock, and the viewer translates through its `t`.
 *
 * This relies on react-i18next internals: `useTranslation` builds the stand-in from
 * `i18n || {}` (wrapped by `createI18nWrapper` from 16.6), which is not public API. Check it
 * first whenever react-i18next is upgraded. A change there fails safe (raw keys, no crash), and
 * the tests that run the real react-i18next without an instance catch it.
 */
function isMissingInstance(result: { readonly ready?: unknown; readonly i18n?: unknown }): boolean {
	return result.ready === false && isEmptyPlainObject(result.i18n);
}

/** The `i18n` of a mock that returns none. */
const NO_FIELDS: HostI18nextInstance = {};

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
 * react-i18next's own `useTranslation`, called unconditionally, so Suspense, `ready` and the
 * `languageChanged` re-render behave exactly as react-i18next defines them. `null` when
 * react-i18next has no instance (neither an `I18nextProvider` nor `initReactI18next`): it then
 * hands back its stand-in, whose `t` returns raw keys, so the viewer shows its built-in English
 * text instead. Only `useTranslation` is read, so a host's module mock that defines nothing else
 * keeps working.
 */
export function useHostI18next(): HostI18next | null {
	const result = useReactI18nextTranslation();
	if (isMissingInstance(result)) {
		return null;
	}
	const { t, i18n } = result as { readonly t: unknown; readonly i18n?: unknown };
	// i18next's `TFunction` accepts every call shape the viewer uses (an options object with
	// `defaultValue` and `{{name}}` params, or a default string); only its overloads differ.
	return {
		t: t as ViewerT,
		i18n: typeof i18n === 'object' && i18n !== null ? (i18n as HostI18nextInstance) : NO_FIELDS,
	};
}

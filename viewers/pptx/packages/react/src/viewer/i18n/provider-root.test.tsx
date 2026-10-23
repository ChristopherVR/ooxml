// @vitest-environment happy-dom
/**
 * A viewer root inside `PptxI18nProvider`: the host owns the language, so the root lists
 * `availableLocales`, hands a pick to `onLocaleChange`, offers the Language tab only when a pick
 * can apply, and never reads, applies or stores a language itself.
 */
import { createInstance } from 'i18next';
import {
	VIEWER_PREFS_STORAGE_KEY,
	readStoredViewerPrefs,
	writeStoredViewerPrefs,
} from 'ooxml-ui/pptx';
import type { LocaleCatalogEntry } from 'ooxml-ui/pptx/i18n';
import { act } from 'react';
import type { ReactNode } from 'react';
import { createRoot } from 'react-dom/client';
import type { Root } from 'react-dom/client';
import * as reactI18next from 'react-i18next';
import { I18nextProvider } from 'react-i18next';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

import { PptxI18nProvider } from './provider';
import { ViewerI18nRoot, useViewerLocaleControls } from './root';
import type { ViewerI18nProps } from './root';
import { useTranslation } from './use-translation';

vi.mock(import('react-i18next'), async (importOriginal) => {
	const actual = await importOriginal();
	return { ...actual, useTranslation: vi.fn(actual.useTranslation) };
});

const SAVED = 'pptx.statusBar.allSaved';
const EN_FR: LocaleCatalogEntry[] = [
	{ code: 'en', label: 'English', nativeLabel: 'English' },
	{ code: 'fr', label: 'French', nativeLabel: 'Francais' },
];

let host: HTMLDivElement;
let root: Root;
let actEnvironment: boolean | undefined;

beforeEach(() => {
	actEnvironment = globalThis.IS_REACT_ACT_ENVIRONMENT;
	globalThis.IS_REACT_ACT_ENVIRONMENT = true;
	localStorage.removeItem(VIEWER_PREFS_STORAGE_KEY);
	vi.mocked(reactI18next.useTranslation).mockClear();
	host = document.createElement('div');
	document.body.append(host);
	root = createRoot(host);
});

afterEach(() => {
	act(() => {
		root.unmount();
	});
	host.remove();
	localStorage.removeItem(VIEWER_PREFS_STORAGE_KEY);
	globalThis.IS_REACT_ACT_ENVIRONMENT = actEnvironment;
});

/** The Language tab state a viewer component sees, plus a French pick. */
function LanguageProbe() {
	const { t } = useTranslation();
	const { localeCode, locales, languageSelectable, selectLocale } = useViewerLocaleControls();
	return (
		<div>
			<output>{t(SAVED)}</output>
			<span data-locale={localeCode} data-selectable={String(languageSelectable)}>
				{locales.map((entry) => entry.code).join(',')}
			</span>
			<button type='button' onClick={() => selectLocale('fr')}>
				pick
			</button>
		</div>
	);
}

const shown = () => host.querySelector<HTMLElement>('[data-locale]')!;

function renderRoot(
	props: ViewerI18nProps,
	{
		locale = 'de',
		wrap = (node) => node,
	}: { locale?: string; wrap?: (node: ReactNode) => ReactNode } = {},
) {
	act(() => {
		root.render(
			wrap(
				<PptxI18nProvider translate={(key) => `host:${key}`} locale={locale}>
					<ViewerI18nRoot {...props}>
						<LanguageProbe />
					</ViewerI18nRoot>
				</PptxI18nProvider>,
			),
		);
	});
}

describe('viewer root inside PptxI18nProvider', () => {
	it('lists availableLocales, selects the provider locale and hands a pick to onLocaleChange', () => {
		const onLocaleChange = vi.fn<(code: string) => void>();
		renderRoot({ availableLocales: EN_FR, onLocaleChange });
		expect(host.querySelector('output')?.textContent).toBe(`host:${SAVED}`);
		expect(shown().dataset.locale).toBe('de');
		expect(shown().textContent).toBe('en,fr');
		expect(shown().dataset.selectable).toBe('true');
		act(() => {
			host.querySelector('button')?.click();
		});
		expect(onLocaleChange).toHaveBeenCalledWith('fr');
		expect(shown().dataset.locale).toBe('de');
		expect(readStoredViewerPrefs().localeCode).toBeUndefined();
	});

	it('selects an offered base language for a regional provider locale', () => {
		renderRoot({ availableLocales: EN_FR, onLocaleChange: () => {} }, { locale: 'fr-CA' });
		expect(shown().dataset.locale).toBe('fr');
		renderRoot({ availableLocales: EN_FR, onLocaleChange: () => {} }, { locale: 'it' });
		expect(shown().dataset.locale).toBe('it');
	});

	it('leaves the Language tab out without onLocaleChange', () => {
		renderRoot({ availableLocales: EN_FR });
		expect(shown().dataset.selectable).toBe('false');
	});

	it('leaves the Language tab out with onLocaleChange but without availableLocales', () => {
		renderRoot({ onLocaleChange: () => {} });
		expect(shown().textContent).toBe('');
		expect(shown().dataset.selectable).toBe('false');
	});

	it('leaves the Language tab out with a single available locale', () => {
		renderRoot({ availableLocales: EN_FR.slice(0, 1), onLocaleChange: () => {} });
		expect(shown().textContent).toBe('en');
		expect(shown().dataset.selectable).toBe('false');
	});

	it('neither applies a stored choice nor changes a host i18next language', async () => {
		const i18n = createInstance();
		await i18n.init({ lng: 'en', resources: { en: { translation: {} } } });
		const changeLanguage = vi.spyOn(i18n, 'changeLanguage');
		writeStoredViewerPrefs({ localeCode: 'fr' });
		expect(readStoredViewerPrefs().localeCode).toBe('fr');
		renderRoot(
			{ defaultLocale: 'es', availableLocales: EN_FR, onLocaleChange: () => {} },
			{ wrap: (node) => <I18nextProvider i18n={i18n}>{node}</I18nextProvider> },
		);
		act(() => {
			host.querySelector('button')?.click();
		});
		expect(changeLanguage).not.toHaveBeenCalled();
		expect(reactI18next.useTranslation).not.toHaveBeenCalled();
		expect(shown().dataset.locale).toBe('de');
		expect(readStoredViewerPrefs().localeCode).toBe('fr');
	});
});

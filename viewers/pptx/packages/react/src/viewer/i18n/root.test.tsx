// @vitest-environment happy-dom
import { createInstance } from 'i18next';
import type { i18n as I18nInstance } from 'i18next';
import { VIEWER_PREFS_STORAGE_KEY, readStoredViewerPrefs } from 'ooxml-ui/pptx';
import { act } from 'react';
import type { ReactNode } from 'react';
import { createRoot } from 'react-dom/client';
import type { Root } from 'react-dom/client';
import { I18nextProvider } from 'react-i18next';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

import { ViewerI18nRoot, useViewerLocaleControls } from './root';
import type { ViewerI18nProps } from './root';
import { useTranslation } from './use-translation';

const SAVED = 'pptx.statusBar.allSaved';

async function hostInstance(): Promise<I18nInstance> {
	const i18n = createInstance();
	await i18n.init({
		lng: 'en',
		fallbackLng: 'en',
		resources: {
			en: { translation: { [SAVED]: 'All saved' } },
			fr: { translation: { [SAVED]: 'Tout est enregistre' } },
			de: { translation: { [SAVED]: 'Alles gespeichert' } },
		},
		interpolation: { escapeValue: false },
	});
	return i18n;
}

/** Shows what a viewer component sees: its text and the Language tab state. */
function Probe() {
	const { t } = useTranslation();
	const { localeCode, locales, selectLocale } = useViewerLocaleControls();
	return (
		<div>
			<output>{[t(SAVED), t('pptx.probe.missing', 'Fallback text')].join('|')}</output>
			<span data-locale={localeCode}>{locales.map((entry) => entry.code).join(',')}</span>
			<button type='button' onClick={() => selectLocale('fr')}>
				pick
			</button>
		</div>
	);
}

describe('viewer translation root (react-i18next source)', () => {
	let host: HTMLDivElement;
	let root: Root;
	let i18n: I18nInstance;
	let actEnvironment: boolean | undefined;

	beforeEach(async () => {
		actEnvironment = globalThis.IS_REACT_ACT_ENVIRONMENT;
		globalThis.IS_REACT_ACT_ENVIRONMENT = true;
		localStorage.removeItem(VIEWER_PREFS_STORAGE_KEY);
		i18n = await hostInstance();
		host = document.createElement('div');
		document.body.append(host);
		root = createRoot(host);
	});

	afterEach(() => {
		act(() => root.unmount());
		host.remove();
		localStorage.removeItem(VIEWER_PREFS_STORAGE_KEY);
		vi.restoreAllMocks();
		globalThis.IS_REACT_ACT_ENVIRONMENT = actEnvironment;
	});

	function render(node: ReactNode) {
		act(() => root.render(<I18nextProvider i18n={i18n}>{node}</I18nextProvider>));
	}
	const text = () => host.querySelector('output')?.textContent;
	const codes = () => host.querySelector('span')?.textContent;
	const selected = () => host.querySelector('span')?.dataset.locale;
	const pick = () => act(() => host.querySelector('button')?.click());
	const viewer = (props: ViewerI18nProps = {}) => (
		<ViewerI18nRoot {...props}>
			<Probe />
		</ViewerI18nRoot>
	);

	it('translates a standalone building block through the host instance and follows it', async () => {
		render(<Probe />);
		expect(text()).toBe('All saved|Fallback text');
		await act(() => i18n.changeLanguage('de'));
		expect(text()).toBe('Alles gespeichert|Fallback text');
	});

	it('lists the registered resources and applies a pick with changeLanguage', async () => {
		const changeLanguage = vi.spyOn(i18n, 'changeLanguage');
		render(viewer());
		expect(codes()).toBe('en,fr,de');
		expect(selected()).toBe('en');
		expect(changeLanguage).not.toHaveBeenCalled();
		await act(async () => pick());
		expect(changeLanguage).toHaveBeenCalledWith('fr');
		expect(text()).toBe('Tout est enregistre|Fallback text');
		expect(selected()).toBe('fr');
		expect(readStoredViewerPrefs().localeCode).toBe('fr');
	});

	it('follows a host language change made outside the viewer', async () => {
		render(viewer());
		await act(() => i18n.changeLanguage('de'));
		expect(text()).toBe('Alles gespeichert|Fallback text');
	});

	it('applies a stored non-English choice once on mount', async () => {
		localStorage.setItem(VIEWER_PREFS_STORAGE_KEY, JSON.stringify({ localeCode: 'de' }));
		const changeLanguage = vi.spyOn(i18n, 'changeLanguage');
		await act(async () => render(viewer()));
		render(viewer({ availableLocales: undefined }));
		expect(changeLanguage).toHaveBeenCalledExactlyOnceWith('de');
		expect(text()).toBe('Alles gespeichert|Fallback text');
		expect(selected()).toBe('de');
	});

	it('applies defaultLocale on mount like a stored choice', async () => {
		const changeLanguage = vi.spyOn(i18n, 'changeLanguage');
		await act(async () => render(viewer({ defaultLocale: 'fr' })));
		expect(changeLanguage).toHaveBeenCalledExactlyOnceWith('fr');
		expect(selected()).toBe('fr');
	});

	it('lets onLocaleChange own the choice: no changeLanguage and nothing persisted', async () => {
		localStorage.setItem(VIEWER_PREFS_STORAGE_KEY, JSON.stringify({ localeCode: 'de' }));
		const changeLanguage = vi.spyOn(i18n, 'changeLanguage');
		const onLocaleChange = vi.fn();
		await act(async () => render(viewer({ onLocaleChange })));
		expect(changeLanguage).not.toHaveBeenCalled();
		pick();
		expect(onLocaleChange).toHaveBeenCalledExactlyOnceWith('fr');
		expect(changeLanguage).not.toHaveBeenCalled();
		expect(selected()).toBe('fr');
		expect(text()).toBe('All saved|Fallback text');
		expect(readStoredViewerPrefs().localeCode).toBe('de');
	});

	it('offers availableLocales instead of the registered resources', () => {
		render(
			viewer({ availableLocales: [{ code: 'es', label: 'Spanish', nativeLabel: 'Espanol' }] }),
		);
		expect(codes()).toBe('es');
	});
});

// @vitest-environment happy-dom
/**
 * `PptxI18nProvider` end to end through the real `PowerPointViewer`, with no i18next instance:
 * the ribbon comes from the host translator (English filling its misses), react-i18next is
 * never called, a `locale` change re-labels the ribbon and an open File > Options, and File >
 * Options shows its Language category only with `onLocaleChange`, which a pick then calls.
 */
import { translationsEn } from 'ooxml-ui/pptx/i18n';
import type { LocaleCatalogEntry } from 'ooxml-ui/pptx/i18n';
import React, { act, createRef } from 'react';
import { createRoot } from 'react-dom/client';
import type { Root } from 'react-dom/client';
import * as reactI18next from 'react-i18next';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

vi.mock(import('react-i18next'), async (importOriginal) => {
	const actual = await importOriginal();
	return { ...actual, useTranslation: vi.fn(actual.useTranslation) };
});

const { PptxHandler } = await import('pptx-viewer-core');
const { PowerPointViewer, PptxI18nProvider } = await import('../index');
type ViewerHandle = import('../index').PowerPointViewerHandle;
type ViewerProps = import('../index').PowerPointViewerProps;
type PptxTranslate = import('../index').PptxTranslate;

const INSERT = 'pptx.ribbon.tab.insert';
const HOME = 'pptx.ribbon.tab.home';
const GENERAL = 'pptx.settings.general';
const translate: PptxTranslate = (key) => (key === INSERT ? 'Einfuegen (host)' : key);

/** A host catalog with a module-level translator: only `locale` tells the viewer to re-label. */
const CATALOGS: Readonly<Record<string, Readonly<Record<string, string>>>> = {
	de: { [INSERT]: 'Einfuegen', [GENERAL]: 'Allgemein' },
	fr: { [INSERT]: 'Insertion', [GENERAL]: 'General (fr)' },
};
let active = 'de';
const catalogTranslate: PptxTranslate = (key) => CATALOGS[active]?.[key];

async function sampleDeck(): Promise<Uint8Array> {
	const { handler, data } = await PptxHandler.create({ initialSlideCount: 1 });
	try {
		return await handler.save(data.slides);
	} finally {
		handler.dispose();
	}
}

let container: HTMLDivElement;
let root: Root;
let warn: ReturnType<typeof vi.spyOn>;
let error: ReturnType<typeof vi.spyOn>;
let actEnvironment: boolean | undefined;

beforeEach(() => {
	actEnvironment = globalThis.IS_REACT_ACT_ENVIRONMENT;
	globalThis.IS_REACT_ACT_ENVIRONMENT = true;
	window.happyDOM?.setViewport({ width: 1600, height: 950 });
	vi.mocked(reactI18next.useTranslation).mockClear();
	warn = vi.spyOn(console, 'warn');
	error = vi.spyOn(console, 'error');
	container = document.createElement('div');
	document.body.appendChild(container);
	root = createRoot(container);
});

afterEach(() => {
	act(() => {
		root.unmount();
	});
	container.remove();
	warn.mockRestore();
	error.mockRestore();
	globalThis.IS_REACT_ACT_ENVIRONMENT = actEnvironment;
});

const ribbonTabs = () =>
	Array.from(container.querySelectorAll('[role="tab"]')).map(
		(tab) => tab.textContent?.trim() ?? '',
	);

const tick = (ms = 20) =>
	act(async () => {
		await new Promise((resolve) => {
			setTimeout(resolve, ms);
		});
	});

type Host = { translate: PptxTranslate; locale: string };

async function mount(
	props: Partial<ViewerProps> = {},
	host: Host = { translate, locale: 'de' },
): Promise<ViewerHandle> {
	const content = await sampleDeck();
	const ref = createRef<ViewerHandle>();
	rerender = async (next: Host) => {
		await act(async () => {
			root.render(
				<PptxI18nProvider translate={next.translate} locale={next.locale}>
					<PowerPointViewer ref={ref} content={content} canEdit {...props} />
				</PptxI18nProvider>,
			);
		});
	};
	await rerender(host);
	for (let attempt = 0; attempt < 80; attempt += 1) {
		if (ref.current && (ref.current.getSlideCount() ?? 0) > 0 && ribbonTabs().length > 0) {
			return ref.current;
		}
		await tick();
	}
	throw new Error('the viewer never finished loading its deck');
}

let rerender: (host: Host) => Promise<void>;

type OptionsDialog = HTMLElement & {
	categories: readonly { id: string; label: string }[];
	updateComplete: Promise<unknown>;
};

/** Open File > Options (the File tab is the first ribbon tab, Options the last nav entry). */
async function openOptions(): Promise<OptionsDialog> {
	await act(async () => {
		container.querySelector<HTMLElement>('[role="tab"]')?.click();
	});
	await tick();
	const entries = document.body.querySelectorAll<HTMLElement>('[data-pptx-backstage-nav-item]');
	await act(async () => {
		entries[entries.length - 1]?.click();
	});
	await tick();
	const dialog = document.body.querySelector<OptionsDialog>('pptx-ui-options-dialog');
	if (!dialog) {
		throw new Error('File > Options did not open');
	}
	await act(async () => {
		await dialog.updateComplete;
	});
	return dialog;
}

/** The category tabs the rendered Options dialog lists in its rail. */
const categoryTabs = (dialog: OptionsDialog) =>
	Array.from(dialog.shadowRoot?.querySelectorAll<HTMLElement>('nav [role="tab"]') ?? []);
const shownCategories = (dialog: OptionsDialog) =>
	categoryTabs(dialog).map((tab) => tab.dataset.category);

const EN_FR: LocaleCatalogEntry[] = [
	{ code: 'en', label: 'English', nativeLabel: 'English' },
	{ code: 'fr', label: 'French', nativeLabel: 'Francais' },
];

const i18nextMessages = (spy: ReturnType<typeof vi.spyOn>) =>
	spy.mock.calls.flat().filter((arg: unknown) => /i18next/i.test(String(arg)));

describe('powerPointViewer inside PptxI18nProvider', () => {
	it('localizes the chrome from the host translator without touching react-i18next', async () => {
		await mount();
		expect(ribbonTabs()).toContain('Einfuegen (host)');
		expect(ribbonTabs()).toContain(translationsEn[HOME]);
		expect(ribbonTabs().filter((label) => label.startsWith('pptx.'))).toStrictEqual([]);
		expect(reactI18next.useTranslation).not.toHaveBeenCalled();
		expect(i18nextMessages(warn)).toStrictEqual([]);
		expect(i18nextMessages(error)).toStrictEqual([]);
	});

	it('re-labels the ribbon and an open File > Options when only locale changes', async () => {
		active = 'de';
		await mount({}, { translate: catalogTranslate, locale: 'de' });
		expect(ribbonTabs()).toContain('Einfuegen');
		const dialog = await openOptions();
		const general = () => dialog.categories.find((entry) => entry.id === 'general')?.label;
		expect(general()).toBe('Allgemein');
		active = 'fr';
		await rerender({ translate: catalogTranslate, locale: 'fr' });
		await act(async () => {
			await dialog.updateComplete;
		});
		expect(ribbonTabs()).toContain('Insertion');
		expect(ribbonTabs()).not.toContain('Einfuegen');
		expect(general()).toBe('General (fr)');
		expect(document.body.querySelector('pptx-ui-options-dialog')).toBe(dialog);
	});

	it('leaves Language out of File > Options without onLocaleChange', async () => {
		await mount({ availableLocales: EN_FR });
		const dialog = await openOptions();
		expect(shownCategories(dialog)).toContain('general');
		expect(shownCategories(dialog)).not.toContain('language');
	});

	it('lists availableLocales and hands a pick to onLocaleChange', async () => {
		const onLocaleChange = vi.fn<(code: string) => void>();
		await mount({ availableLocales: EN_FR, onLocaleChange });
		const dialog = await openOptions();
		expect(shownCategories(dialog)).toContain('language');
		await act(async () => {
			categoryTabs(dialog)
				.find((tab) => tab.dataset.category === 'language')
				?.click();
		});
		await tick();
		const choices = Array.from(
			document.body.querySelectorAll<HTMLButtonElement>('[slot="block-language"] button'),
		);
		expect(choices.map((choice) => choice.textContent)).toStrictEqual([
			'English',
			'FrancaisFrench',
		]);
		await act(async () => {
			choices[1]?.click();
		});
		expect(onLocaleChange).toHaveBeenCalledWith('fr');
		expect(reactI18next.useTranslation).not.toHaveBeenCalled();
	});
});

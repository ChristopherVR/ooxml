// @vitest-environment happy-dom
// @vitest-fresh-modules
/**
 * `pptx-react-viewer/no-i18next`: the main entry's API without react-i18next. Outside a
 * `PptxI18nProvider` its components show the built-in English text, even under an
 * `I18nextProvider` (proof that they never read react-i18next); inside one they translate
 * through it. The main entry keeps reading react-i18next outside a provider, and every entry
 * falls back to the built-in English text while react-i18next has no configured instance.
 *
 * The test setup registers react-i18next for every spec, as importing the main entry does, so the
 * no-i18next cases load a fresh module graph from `./no-i18next` alone, which leaves the
 * registration slot empty exactly as a host bundle of that entry does.
 */
import { createInstance } from 'i18next';
import type { i18n as I18nInstance } from 'i18next';
import { VIEWER_PREFS_STORAGE_KEY, writeStoredViewerPrefs } from 'ooxml-ui/pptx';
import { translationsEn } from 'ooxml-ui/pptx/i18n';
import { act } from 'react';
import type { ComponentType, ReactNode } from 'react';
import { createRoot } from 'react-dom/client';
import type { Root } from 'react-dom/client';
import { I18nextProvider, getI18n, setI18n } from 'react-i18next';
import { afterAll, afterEach, beforeAll, beforeEach, describe, expect, it, vi } from 'vitest';

import type { ToolbarProps } from './public-api';

const LOADING = 'pptx.viewer.loading';
const HOME = 'pptx.ribbon.tab.home';
const SAVED = 'pptx.statusBar.allSaved';
const FROM_I18NEXT = 'From i18next';

type Entry = typeof import('./no-i18next');

let host: HTMLDivElement;
let root: Root;
let i18n: I18nInstance;
let actEnvironment: boolean | undefined;

beforeAll(async () => {
	// Every key the cases render, so a component that reads react-i18next shows FROM_I18NEXT.
	i18n = createInstance();
	await i18n.init({
		lng: 'en',
		resources: {
			en: {
				translation: { [LOADING]: FROM_I18NEXT, [HOME]: FROM_I18NEXT, [SAVED]: FROM_I18NEXT },
			},
		},
		interpolation: { escapeValue: false },
	});
});

beforeEach(() => {
	actEnvironment = globalThis.IS_REACT_ACT_ENVIRONMENT;
	globalThis.IS_REACT_ACT_ENVIRONMENT = true;
	localStorage.removeItem(VIEWER_PREFS_STORAGE_KEY);
	host = document.createElement('div');
	document.body.append(host);
	root = createRoot(host);
});

afterEach(() => {
	act(() => root.unmount());
	host.remove();
	localStorage.removeItem(VIEWER_PREFS_STORAGE_KEY);
	globalThis.IS_REACT_ACT_ENVIRONMENT = actEnvironment;
});

/** Renders under an `I18nextProvider`, so only a component that never reads react-i18next misses it. */
function render(node: ReactNode): void {
	act(() => root.render(<I18nextProvider i18n={i18n}>{node}</I18nextProvider>));
}

const text = () => document.body.textContent ?? '';

/** Enough of the toolbar's props to render its tab strip; no handler runs while rendering. */
function toolbarProps(): ToolbarProps {
	return {
		mode: 'edit',
		canEdit: true,
		isNarrowViewport: false,
		isSidebarCollapsed: false,
		isInspectorPaneOpen: false,
		isCompactToolbarOpen: false,
		toolbarSection: 'home',
		scale: 1,
		canUndo: false,
		canRedo: false,
		findReplaceOpen: false,
		selectedElement: null,
		editTemplateMode: false,
		newShapeType: 'rect',
		activeTool: 'select',
		drawingColor: '#000000',
		drawingWidth: 2,
		clipboardPayload: null,
		spellCheckEnabled: false,
		showGrid: false,
		showRulers: false,
		snapToGrid: false,
		snapToShape: false,
		layoutOptions: [],
		customShows: [],
		activeCustomShowId: null,
		isCurrentSlideInActiveShow: false,
		isThemeEditorOpen: false,
		isThemeGalleryOpen: false,
		isOverflowMenuOpen: false,
		onSetToolbarSection: vi.fn<() => void>(),
	} as unknown as ToolbarProps;
}

describe('pptx-react-viewer/no-i18next', () => {
	let entry: Entry;
	let i18nRoot: typeof import('./viewer/i18n/root');
	let translation: typeof import('./viewer/i18n/use-translation');
	let slot: typeof import('./viewer/i18n/host-i18next');

	beforeAll(async () => {
		vi.resetModules();
		entry = await import('./no-i18next');
		// The same fresh graph the entry loaded: these resolve to the modules it imports.
		i18nRoot = await import('./viewer/i18n/root');
		translation = await import('./viewer/i18n/use-translation');
		slot = await import('./viewer/i18n/host-i18next');
	}, 60_000);

	afterAll(() => {
		vi.resetModules();
	});

	it('registers no react-i18next source', () => {
		expect(slot.registeredHostI18next()).toBeNull();
	});

	it('shows the built-in English text in a viewer without a provider', () => {
		render(<entry.PowerPointViewer content={null} />);
		expect(text()).toContain(translationsEn[LOADING]);
		expect(text()).not.toContain(FROM_I18NEXT);
	});

	it('shows the built-in English text in a standalone Toolbar without a provider', () => {
		const Toolbar = entry.Toolbar as ComponentType<ToolbarProps>;
		render(<Toolbar {...toolbarProps()} />);
		expect(text()).toContain(translationsEn[HOME]);
		expect(text()).not.toContain(FROM_I18NEXT);
	});

	it('translates the viewer and a standalone block through PptxI18nProvider', () => {
		const translate = (key: string) => `host:${key}`;
		const Toolbar = entry.Toolbar as ComponentType<ToolbarProps>;
		render(
			<entry.PptxI18nProvider translate={translate} locale='de'>
				<entry.PowerPointViewer content={null} />
				<Toolbar {...toolbarProps()} />
			</entry.PptxI18nProvider>,
		);
		expect(text()).toContain(`host:${LOADING}`);
		expect(text()).toContain(`host:${HOME}`);
		expect(text()).not.toContain(FROM_I18NEXT);
	});

	it('leaves the Language tab out and neither reads nor stores a language', () => {
		writeStoredViewerPrefs({ localeCode: 'fr' });
		const onLocaleChange = vi.fn<(code: string) => void>();
		function Probe() {
			const { t, locale } = translation.useTranslation();
			const controls = i18nRoot.useViewerLocaleControls();
			return (
				<output data-locale={locale} data-selectable={String(controls.languageSelectable)}>
					{`${t(SAVED)}|${controls.localeCode}|${controls.locales.length}`}
				</output>
			);
		}
		render(
			<i18nRoot.ViewerI18nRoot
				defaultLocale='de'
				availableLocales={[
					{ code: 'en', label: 'English', nativeLabel: 'English' },
					{ code: 'fr', label: 'French', nativeLabel: 'Francais' },
				]}
				onLocaleChange={onLocaleChange}
			>
				<Probe />
			</i18nRoot.ViewerI18nRoot>,
		);
		const output = host.querySelector('output')!;
		expect(output.textContent).toBe(`${translationsEn[SAVED]}|en|0`);
		expect(output.dataset.locale).toBe('en');
		expect(output.dataset.selectable).toBe('false');
		expect(onLocaleChange).not.toHaveBeenCalled();
	});

	it('keeps reading the slot once per mounted component', () => {
		// A source registered after a component mounted (another entry loaded later) only reaches
		// components mounted after it, so a mounted component never changes its hook calls.
		function Probe({ pass }: { pass: number }) {
			const { t } = translation.useTranslation();
			return <output>{`${pass}:${t(SAVED)}`}</output>;
		}
		render(<Probe pass={1} />);
		const useFromI18next = vi.fn(() => ({
			t: () => FROM_I18NEXT,
			i18n: { language: 'en', options: {}, changeLanguage: () => {} },
		}));
		slot.registerHostI18next(useFromI18next);
		render(<Probe pass={2} />);
		expect(host.querySelector('output')?.textContent).toBe(`2:${translationsEn[SAVED]}`);
		expect(useFromI18next).not.toHaveBeenCalled();
	});
});

describe('pptx-react-viewer (main entry)', () => {
	let main: typeof import('./index');
	let slot: typeof import('./viewer/i18n/host-i18next');
	let source: typeof import('./viewer/i18n/react-i18next-source');

	beforeAll(async () => {
		vi.resetModules();
		main = await import('./index');
		slot = await import('./viewer/i18n/host-i18next');
		source = await import('./viewer/i18n/react-i18next-source');
	}, 60_000);

	afterAll(() => {
		vi.resetModules();
	});

	it('registers react-i18next when it loads', () => {
		expect(slot.registeredHostI18next()).toBe(source.useHostI18next);
	});

	it('exports the same API as pptx-react-viewer/no-i18next', async () => {
		const noI18next = await import('./no-i18next');
		expect(Object.keys(noI18next).toSorted()).toStrictEqual(Object.keys(main).toSorted());
		expect(noI18next.PowerPointViewer).toBe(main.PowerPointViewer);
		expect(noI18next.PptxI18nProvider).toBe(main.PptxI18nProvider);
	});

	it('reads react-i18next in a viewer and a standalone Toolbar without a provider', () => {
		const Toolbar = main.Toolbar as ComponentType<ToolbarProps>;
		render(
			<>
				<main.PowerPointViewer content={null} />
				<Toolbar {...toolbarProps()} />
			</>,
		);
		expect(text()).not.toContain(translationsEn[LOADING]);
		expect(text()).not.toContain(translationsEn[HOME]);
		expect(text().split(FROM_I18NEXT).length - 1).toBeGreaterThanOrEqual(2);
	});

	it('keeps the first registration when an entry registers again', async () => {
		await import('./viewer/index');
		await import('./internals');
		slot.registerHostI18next(() => {
			throw new Error('a second registration must not replace the first');
		});
		expect(slot.registeredHostI18next()).toBe(source.useHostI18next);
		render(<main.PowerPointViewer content={null} />);
		expect(text()).toContain(FROM_I18NEXT);
	});
});

/**
 * The registration is app-wide, so once any module loads the main entry, components from
 * `./no-i18next` read react-i18next too. A host of that entry never configured react-i18next:
 * with no instance (no `I18nextProvider`, no `initReactI18next`) the viewer must keep its
 * built-in English text rather than throw or show raw keys.
 */
describe('after the main entry loads, with react-i18next not configured', () => {
	let noI18next: Entry;
	let main: typeof import('./index');
	let slot: typeof import('./viewer/i18n/host-i18next');
	let previousInstance: I18nInstance | undefined;

	/** Renders with no `I18nextProvider`, so react-i18next only finds its process-wide instance. */
	function renderBare(node: ReactNode): void {
		act(() => root.render(node));
	}

	beforeAll(async () => {
		vi.resetModules();
		noI18next = await import('./no-i18next');
		// The same graph: the stray import registers react-i18next for every entry's components.
		main = await import('./index');
		slot = await import('./viewer/i18n/host-i18next');
		previousInstance = getI18n() as I18nInstance | undefined;
	}, 60_000);

	beforeEach(() => {
		setI18n(undefined as unknown as I18nInstance);
		// react-i18next reports the missing instance once; keep the run quiet.
		vi.spyOn(console, 'warn').mockImplementation(() => {});
	});

	afterEach(() => {
		vi.restoreAllMocks();
	});

	afterAll(() => {
		setI18n(previousInstance as I18nInstance);
		vi.resetModules();
	});

	it('has react-i18next registered for the no-i18next components too', () => {
		expect(slot.registeredHostI18next()).not.toBeNull();
	});

	it('shows the built-in English text in a no-i18next viewer and Toolbar', () => {
		const Toolbar = noI18next.Toolbar as ComponentType<ToolbarProps>;
		expect(() =>
			renderBare(
				<>
					<noI18next.PowerPointViewer content={null} />
					<Toolbar {...toolbarProps()} />
				</>,
			),
		).not.toThrow();
		expect(text()).toContain(translationsEn[LOADING]);
		expect(text()).toContain(translationsEn[HOME]);
		expect(text()).not.toContain('pptx.');
	});

	it('shows the built-in English text in a main-entry viewer instead of throwing', () => {
		const Toolbar = main.Toolbar as ComponentType<ToolbarProps>;
		expect(() =>
			renderBare(
				<>
					<main.PowerPointViewer content={null} />
					<Toolbar {...toolbarProps()} />
				</>,
			),
		).not.toThrow();
		expect(text()).toContain(translationsEn[LOADING]);
		expect(text()).toContain(translationsEn[HOME]);
		expect(text()).not.toContain('pptx.');
	});

	it('still translates through react-i18next once it is configured', () => {
		// What `initReactI18next` does: the process-wide instance, with no `I18nextProvider`.
		setI18n(i18n);
		const Toolbar = noI18next.Toolbar as ComponentType<ToolbarProps>;
		renderBare(
			<>
				<noI18next.PowerPointViewer content={null} />
				<Toolbar {...toolbarProps()} />
			</>,
		);
		expect(text()).not.toContain(translationsEn[LOADING]);
		expect(text()).not.toContain(translationsEn[HOME]);
		expect(text().split(FROM_I18NEXT).length - 1).toBeGreaterThanOrEqual(2);
	});
});

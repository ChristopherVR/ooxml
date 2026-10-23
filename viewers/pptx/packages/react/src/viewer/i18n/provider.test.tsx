// @vitest-environment happy-dom
/**
 * `PptxI18nProvider`: a host translates the viewer with its own i18n library. Covers the
 * fallback chain behind a host miss, `{{name}}` interpolation, params reaching the host
 * unchanged, a stable `t`, which context wins, and when react-i18next is called.
 */
import { createInstance } from 'i18next';
import { keyToLabel, translationsEn } from 'ooxml-ui/pptx/i18n';
import { act } from 'react';
import type { ReactNode } from 'react';
import { createRoot } from 'react-dom/client';
import type { Root } from 'react-dom/client';
import * as reactI18next from 'react-i18next';
import { I18nextProvider } from 'react-i18next';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

import { translationsFr } from '../../../../locales/src';
import { PptxI18nProvider, SlideTemplateGalleryDialog } from '../../index';
import type { PptxTranslate } from '../../index';
import { ViewerI18nRoot } from './root';
import type { ViewerT } from './types';
import { useTranslation } from './use-translation';

vi.mock(import('react-i18next'), async (importOriginal) => {
	const actual = await importOriginal();
	return { ...actual, useTranslation: vi.fn(actual.useTranslation) };
});

const SAVED = 'pptx.statusBar.allSaved';
const INSERT = 'pptx.ribbon.tab.insert';
const USERS = 'pptx.collaboration.userCount';
const UNKNOWN = 'pptx.probe.notInTheDictionary';

let host: HTMLDivElement;
let root: Root;
let warn: ReturnType<typeof vi.spyOn>;
let error: ReturnType<typeof vi.spyOn>;
let actEnvironment: boolean | undefined;

beforeEach(() => {
	actEnvironment = globalThis.IS_REACT_ACT_ENVIRONMENT;
	globalThis.IS_REACT_ACT_ENVIRONMENT = true;
	vi.mocked(reactI18next.useTranslation).mockClear();
	warn = vi.spyOn(console, 'warn');
	error = vi.spyOn(console, 'error');
	host = document.createElement('div');
	document.body.append(host);
	root = createRoot(host);
});

afterEach(() => {
	act(() => {
		root.unmount();
	});
	host.remove();
	warn.mockRestore();
	error.mockRestore();
	globalThis.IS_REACT_ACT_ENVIRONMENT = actEnvironment;
});

function render(node: ReactNode): void {
	act(() => {
		root.render(node);
	});
}

/** Renders what `t` returns for each call, one `<li>` per call, and hands `t` out. */
function Calls({ calls, seen }: { calls: ((t: ViewerT) => string)[]; seen?: ViewerT[] }) {
	const { t, locale } = useTranslation();
	seen?.push(t);
	return (
		<ul data-locale={locale}>
			{calls.map((call, index) => (
				<li key={index}>{call(t)}</li>
			))}
		</ul>
	);
}

const texts = () => Array.from(host.querySelectorAll('li'), (item) => item.textContent);

describe('pptxI18nProvider', () => {
	it('translates through the host function and reports its locale', () => {
		render(
			<PptxI18nProvider translate={(key) => `host:${key}`} locale='fr-CA'>
				<Calls calls={[(t) => t(SAVED)]} />
			</PptxI18nProvider>,
		);
		expect(texts()).toStrictEqual([`host:${SAVED}`]);
		expect(host.querySelector('ul')?.dataset.locale).toBe('fr-CA');
	});

	it('falls back to English, then defaultValue, then a key label on a host miss', () => {
		const misses: PptxTranslate = (key) => (key === SAVED ? '' : key);
		render(
			<PptxI18nProvider translate={misses} locale='de'>
				<Calls
					calls={[
						(t) => t(SAVED),
						(t) => t(INSERT, { defaultValue: 'Not this' }),
						(t) => t(INSERT, 'Not this either'),
						(t) => t(UNKNOWN, { defaultValue: 'Object {{n}}', n: 2 }),
						(t) => t(UNKNOWN, 'String default'),
						(t) => t(UNKNOWN),
					]}
				/>
			</PptxI18nProvider>,
		);
		expect(texts()).toStrictEqual([
			translationsEn[SAVED],
			translationsEn[INSERT],
			translationsEn[INSERT],
			'Object 2',
			'String default',
			keyToLabel(UNKNOWN),
		]);
		expect(texts().filter((text) => text?.startsWith('pptx.'))).toStrictEqual([]);
	});

	it('treats undefined, null and a throw from the host as a miss', () => {
		const translate: PptxTranslate = (key) => {
			if (key === SAVED) {
				return undefined;
			}
			if (key === INSERT) {
				return null;
			}
			if (key === UNKNOWN) {
				throw new Error('catalog not loaded');
			}
			return 'Kept';
		};
		render(
			<PptxI18nProvider translate={translate} locale='de'>
				<Calls
					calls={[
						(t) => t(SAVED),
						(t) => t(INSERT),
						(t) => t(UNKNOWN, 'After a throw'),
						(t) => t('pptx.ribbon.tab.home'),
					]}
				/>
			</PptxI18nProvider>,
		);
		expect(texts()).toStrictEqual([
			translationsEn[SAVED],
			translationsEn[INSERT],
			'After a throw',
			'Kept',
		]);
	});

	it('fills {{name}} placeholders of the English fallback from the params', () => {
		render(
			<PptxI18nProvider translate={(k) => k} locale='en'>
				<Calls calls={[(t) => t(USERS, { count: 7 })]} />
			</PptxI18nProvider>,
		);
		expect(translationsEn[USERS]).toContain('{{count}}');
		expect(texts()[0]).toBe(translationsEn[USERS]!.replace('{{count}}', '7'));
	});

	it('fills {{name}} placeholders left in the host text, so a shipped dictionary works as is', () => {
		const dictionary: Readonly<Record<string, string>> = translationsFr;
		render(
			<PptxI18nProvider translate={(key) => dictionary[key] ?? key} locale='fr'>
				<Calls calls={[(t) => t(USERS, { count: 3 })]} />
			</PptxI18nProvider>,
		);
		expect(dictionary[USERS]).toBe('{{count}} utilisateur(s)');
		expect(texts()).toStrictEqual(['3 utilisateur(s)']);
	});

	it('passes the params to the host unchanged', () => {
		const translate = vi.fn<PptxTranslate>((key) => key);
		const params = { defaultValue: 'Default', count: 0, name: 'Deck' };
		render(
			<PptxI18nProvider translate={translate} locale='en'>
				<Calls calls={[(t) => t(SAVED, params), (t) => t(SAVED, 'Shorthand'), (t) => t(SAVED)]} />
			</PptxI18nProvider>,
		);
		expect(translate.mock.calls[0]?.[1]).toBe(params);
		expect(translate.mock.calls[1]).toStrictEqual([SAVED, { defaultValue: 'Shorthand' }]);
		expect(translate.mock.calls[2]).toStrictEqual([SAVED]);
	});

	it('keeps t while translate and locale stay the same, and replaces it when either changes', () => {
		const seen: ViewerT[] = [];
		const translate: PptxTranslate = (key) => key;
		const tree = (fn: PptxTranslate, locale: string, label: string) => (
			<PptxI18nProvider translate={fn} locale={locale}>
				<Calls calls={[() => label]} seen={seen} />
			</PptxI18nProvider>
		);
		render(tree(translate, 'en', 'first'));
		render(tree(translate, 'en', 'host re-render'));
		expect(texts()).toStrictEqual(['host re-render']);
		expect(seen).toHaveLength(2);
		expect(seen[1]).toBe(seen[0]);
		render(tree(translate, 'fr', 'new locale'));
		expect(seen[2]).not.toBe(seen[1]);
		render(tree((key) => key, 'fr', 'new translate'));
		expect(seen[3]).not.toBe(seen[2]);
	});

	it('translates a standalone building block', () => {
		render(
			<PptxI18nProvider translate={(key) => `host:${key}`} locale='en'>
				<SlideTemplateGalleryDialog isOpen onClose={() => {}} onInsert={() => {}} />
			</PptxI18nProvider>,
		);
		expect(document.body.textContent).toContain('host:pptx.slideTemplates.galleryTitle');
	});

	it('never calls react-i18next and warns nothing without an i18next instance', () => {
		render(
			<PptxI18nProvider translate={(key) => key} locale='en'>
				<ViewerI18nRoot>
					<Calls calls={[(t) => t(SAVED)]} />
				</ViewerI18nRoot>
				<SlideTemplateGalleryDialog isOpen onClose={() => {}} onInsert={() => {}} />
			</PptxI18nProvider>,
		);
		expect(texts()).toStrictEqual([translationsEn[SAVED]]);
		expect(reactI18next.useTranslation).not.toHaveBeenCalled();
		expect(warn).not.toHaveBeenCalled();
		expect(error).not.toHaveBeenCalled();
	});

	it('still reads react-i18next without a provider, once per root and per standalone block', async () => {
		const i18n = createInstance();
		await i18n.init({ lng: 'en', resources: { en: { translation: { [SAVED]: 'From i18next' } } } });
		render(
			<I18nextProvider i18n={i18n}>
				<ViewerI18nRoot>
					<Calls calls={[(t) => t(SAVED)]} />
					<Calls calls={[(t) => t(SAVED)]} />
				</ViewerI18nRoot>
				<Calls calls={[(t) => t(SAVED)]} />
			</I18nextProvider>,
		);
		expect(texts()).toStrictEqual(['From i18next', 'From i18next', 'From i18next']);
		// The root and the standalone `Calls`; the two `Calls` inside the root read the root.
		expect(reactI18next.useTranslation).toHaveBeenCalledTimes(2);
	});

	it('ignores a provider inside a viewer root, while a standalone block reads its nearest', () => {
		const outer: PptxTranslate = (key) => `outer:${key}`;
		const inner: PptxTranslate = (key) => `inner:${key}`;
		render(
			<PptxI18nProvider translate={outer} locale='de'>
				<ViewerI18nRoot>
					<PptxI18nProvider translate={inner} locale='fr'>
						<Calls calls={[(t) => t(SAVED)]} />
					</PptxI18nProvider>
				</ViewerI18nRoot>
				<PptxI18nProvider translate={inner} locale='fr'>
					<Calls calls={[(t) => t(SAVED)]} />
				</PptxI18nProvider>
			</PptxI18nProvider>,
		);
		expect(texts()).toStrictEqual([`outer:${SAVED}`, `inner:${SAVED}`]);
		const locales = Array.from(host.querySelectorAll('ul'), (list) => list.dataset.locale);
		expect(locales).toStrictEqual(['de', 'fr']);
	});
});

// @vitest-environment happy-dom
/**
 * Host tests that mock react-i18next with only `useTranslation`, as react-i18next's testing guide
 * shows: whatever `i18n` the mock returns (none, an empty object, a Proxy auto-mock, one without
 * `changeLanguage`), the viewer translates through the mock's `t` and reads nothing else from the
 * module. Only react-i18next's own stand-in for a missing instance gives the built-in English text.
 */
import { act } from 'react';
import { createRoot } from 'react-dom/client';
import type { Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

import { ViewerI18nRoot, useViewerLocaleControls } from './root';
import { useTranslation } from './use-translation';

const mocked = vi.hoisted(() => ({ result: {} as Record<string, unknown> }));

vi.mock(import('react-i18next'), () => ({
	useTranslation: () => mocked.result,
}));

const t = (key: string) => `M:${key}`;

let host: HTMLDivElement;
let root: Root;
let actEnvironment: boolean | undefined;

beforeEach(() => {
	actEnvironment = globalThis.IS_REACT_ACT_ENVIRONMENT;
	globalThis.IS_REACT_ACT_ENVIRONMENT = true;
	host = document.createElement('div');
	document.body.append(host);
	root = createRoot(host);
});

afterEach(() => {
	act(() => root.unmount());
	host.remove();
	globalThis.IS_REACT_ACT_ENVIRONMENT = actEnvironment;
});

function Probe() {
	const { t: translate, locale } = useTranslation();
	return <output>{`${translate('pptx.statusBar.allSaved')}|${locale}`}</output>;
}

/** The Language tab state a viewer root provides. */
function Controls() {
	const { localeCode, locales, languageSelectable } = useViewerLocaleControls();
	const codes = locales.map((entry) => entry.code).join(',');
	return <span>{`|${localeCode}|${codes}|${languageSelectable}`}</span>;
}

/** What a standalone block shows, then what a viewer root shows. */
function renderBlockAndRoot(result: Record<string, unknown>): readonly [string, string] {
	mocked.result = result;
	act(() => root.render(<Probe />));
	const block = host.textContent ?? '';
	act(() =>
		root.render(
			<ViewerI18nRoot>
				<Probe />
				<Controls />
			</ViewerI18nRoot>,
		),
	);
	return [block, host.textContent ?? ''];
}

/** A deep auto-mock: every field is a function (`then` excepted, so it is not a thenable). */
const autoMock = new Proxy({}, { get: (_target, key) => (key === 'then' ? undefined : () => {}) });

describe('a host mock of react-i18next with only useTranslation', () => {
	it('translates through a mocked i18n without changeLanguage', () => {
		expect(renderBlockAndRoot({ t, i18n: { language: 'de' } })).toStrictEqual([
			'M:pptx.statusBar.allSaved|de',
			'M:pptx.statusBar.allSaved|de|en|en|true',
		]);
	});

	it.each([
		['no i18n', { t }],
		['an empty i18n', { t, i18n: {} }],
		['a Proxy auto-mock for i18n', { t, i18n: autoMock }],
	])('translates through a mock that returns %s', (_name, result) => {
		expect(renderBlockAndRoot(result)).toStrictEqual([
			'M:pptx.statusBar.allSaved|en',
			'M:pptx.statusBar.allSaved|en|en|en|true',
		]);
	});

	it("shows the built-in English text for react-i18next's own stand-in", () => {
		expect(renderBlockAndRoot({ t, i18n: {}, ready: false })).toStrictEqual([
			'All saved|en',
			'All saved|en|en||false',
		]);
	});
});

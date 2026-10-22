/**
 * The theme colour grid and the Quick Styles gallery name their swatches through the provided
 * translator (tooltip and accessible name), not in hard-coded English, and Quick Styles follows a
 * language change.
 */
import type { PptxElement } from 'pptx-viewer-core';
import { translationsDe } from 'pptx-viewer-locales';
import { flushSync, mount, unmount } from 'svelte';
import { afterEach, describe, expect, it } from 'vitest';

import { I18N_CONTEXT_KEY } from '../../../i18n/context';
import { interpolate, translate } from '../../../i18n/translator';
import type { Translator } from '../../../i18n/translator';
import { EditorState } from '../../editor/editor-state.svelte';
import QuickStylesGallery from './QuickStylesGallery.svelte';
import ThemeColorSwatchGrid from './ThemeColorSwatchGrid.svelte';

const OFFICE_THEME = {
	dk1: '#000000',
	lt1: '#FFFFFF',
	dk2: '#44546A',
	lt2: '#E7E6E6',
	accent1: '#4472C4',
	accent2: '#ED7D31',
	accent3: '#A5A5A5',
	accent4: '#FFC000',
	accent5: '#5B9BD5',
	accent6: '#70AD47',
	bg1: '#FFFFFF',
	tx1: '#000000',
	bg2: '#E7E6E6',
	tx2: '#44546A',
};

const german: Translator = (key, params) => {
	const message = (translationsDe as Record<string, string>)[key];
	return message === undefined ? translate('en', key, params) : interpolate(message, params);
};

let cleanup: (() => void) | undefined;
afterEach(() => {
	cleanup?.();
	cleanup = undefined;
});

function render(
	component: Parameters<typeof mount>[0],
	props: Record<string, unknown>,
	translator: Translator = german,
) {
	const target = document.createElement('div');
	document.body.appendChild(target);
	const instance = mount(component, {
		target,
		props,
		context: new Map([[I18N_CONTEXT_KEY, translator]]),
	});
	flushSync();
	cleanup = () => {
		unmount(instance);
		target.remove();
	};
	return target;
}

describe('swatch and style names follow the locale', () => {
	it('names theme colour swatches with the host dictionary', () => {
		const target = render(ThemeColorSwatchGrid, { themeColorMap: OFFICE_THEME, onpick: () => {} });
		// The shared grid draws its swatches and heading in its shadow root.
		const shadow = target.querySelector('pptx-ui-color-grid')!.shadowRoot!;
		const titles = [...shadow.querySelectorAll('button')].map((button) => button.title);
		expect(titles).toContain('Akzent 1, heller 80 %');
		expect(titles.some((title) => title.includes('Accent'))).toBeFalsy();
		expect(shadow.querySelector('.heading')!.textContent).toBe('Designfarben');
	});

	it('names Quick Styles presets with the host dictionary, and relabels on a language change', () => {
		const el = {
			type: 'shape',
			id: 's1',
			x: 0,
			y: 0,
			width: 100,
			height: 50,
			shapeType: 'rect',
			shapeStyle: {},
		} as PptxElement;
		const editor = new EditorState({ getCurrent: () => 0, getHandler: () => null });
		editor.setSlides([{ id: 's1', rId: 'rId1', slideNumber: 1, elements: [el] }]);
		// The viewer's translator reads its locale state on every call; this one does the same.
		const locale = $state({ current: 'de' });
		const t: Translator = (key, params) =>
			locale.current === 'de' ? german(key, params) : translate('en', key, params);
		const target = render(QuickStylesGallery, { editor, el }, t);
		// The tiles are the shared gallery's, inside its element.
		const first = () => target.querySelector('pptx-ui-ribbon-gallery [data-gallery-item]')!;
		expect(first().getAttribute('aria-label')).toBe('Blaue Füllung');
		expect(first().getAttribute('title')).toBe('Blaue Füllung');
		locale.current = 'en';
		flushSync();
		expect(first().getAttribute('aria-label')).toBe('Blue Fill');
		expect(first().getAttribute('title')).toBe('Blue Fill');
	});
});

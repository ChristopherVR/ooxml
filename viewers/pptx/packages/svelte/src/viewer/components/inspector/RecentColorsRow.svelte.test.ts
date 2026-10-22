import { translationsDe } from 'pptx-viewer-locales';
import { flushSync, mount, unmount } from 'svelte';
import { afterEach, describe, expect, it, vi } from 'vitest';

import { I18N_CONTEXT_KEY } from '../../../i18n/context';
import { interpolate, translate } from '../../../i18n/translator';
import type { Translator } from '../../../i18n/translator';
import RecentColorsRow from './RecentColorsRow.svelte';

const german: Translator = (key, params) => {
	const message = (translationsDe as Record<string, string>)[key];
	return message === undefined ? translate('en', key, params) : interpolate(message, params);
};

let cleanup: (() => void) | undefined;
afterEach(() => {
	cleanup?.();
	cleanup = undefined;
});

function render(props: Record<string, unknown>, translator?: Translator) {
	const target = document.createElement('div');
	document.body.appendChild(target);
	const instance = mount(RecentColorsRow, {
		target,
		props: { onselect: vi.fn(), ...props },
		...(translator ? { context: new Map([[I18N_CONTEXT_KEY, translator]]) } : {}),
	});
	flushSync();
	cleanup = () => {
		unmount(instance);
		target.remove();
	};
	const row = () => target.querySelector<HTMLElement>('[data-testid="pptx-color-recent"]');
	const swatches = () => [
		...(row()?.shadowRoot?.querySelectorAll<HTMLButtonElement>('button') ?? []),
	];
	return { row, swatches };
}

describe('recentColorsRow', () => {
	it('renders nothing for an empty list', () => {
		expect(render({ colors: [] }).row()).toBeNull();
	});

	it('is the shared grid with only Recent Colors, applying a pick through onselect', () => {
		const onselect = vi.fn();
		const { row, swatches } = render({ colors: ['#112233', '#445566'], onselect });
		expect(row()?.localName).toBe('pptx-ui-color-grid');
		expect(row()?.getAttribute('aria-label')).toBe('Recent Colors');
		// The cross-binding row contract: the accessible name says which list the swatch came from.
		expect(swatches().map((swatch) => swatch.getAttribute('aria-label'))).toStrictEqual([
			'Recent #112233',
			'Recent #445566',
		]);
		swatches()[1]!.click();
		expect(onselect).toHaveBeenCalledWith('#445566');
	});

	it('names the row and its swatches in the viewer language', () => {
		const { row, swatches } = render({ colors: ['#112233'] }, german);
		expect(row()?.getAttribute('aria-label')).toBe('Zuletzt verwendete Farben');
		expect(swatches()[0]!.getAttribute('aria-label')).toBe('Zuletzt verwendete Farbe #112233');
	});

	it('disables every swatch while the owning picker is disabled', () => {
		const { swatches } = render({ colors: ['#112233'], disabled: true });
		expect(swatches()[0]!.disabled).toBeTruthy();
	});
});

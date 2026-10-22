// @vitest-environment happy-dom
/**
 * B6 (wave-4): recent colours. `RecentColorsRow` is the row every inspector
 * colour picker shows under itself, drawn by the shared `pptx-ui-color-grid`
 * with only its Recent Colors section.
 */
import { translationsEn } from 'ooxml-ui/pptx/i18n';
import React, { act } from 'react';
import { createRoot } from 'react-dom/client';
import type { Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

import { RecentColorsProvider } from './RecentColorsContext';
import { RecentColorsRow } from './RecentColorsRow';

// The English dictionary, so the assertions read the names a user sees.
vi.mock(import('../../i18n/use-translation'), () => ({
	useTranslation: () => ({
		t: (key: string, params?: Record<string, string>) =>
			(translationsEn[key] ?? key).replace(
				/\{\{(\w+)\}\}/gu,
				(_, name: string) => params?.[name] ?? '',
			),
	}),
}));

let container: HTMLDivElement;
let root: Root;

const row = () => container.querySelector('[data-testid="pptx-color-recent"]');
const swatches = () => [
	...(row()?.shadowRoot?.querySelectorAll<HTMLButtonElement>('button[data-source="recent"]') ?? []),
];

beforeEach(() => {
	globalThis.IS_REACT_ACT_ENVIRONMENT = true;
	container = document.createElement('div');
	document.body.appendChild(container);
	root = createRoot(container);
});

afterEach(() => {
	act(() => root.unmount());
	container.remove();
});

function mount(
	recentColors: string[],
	props: { disabled?: boolean } = {},
	pushColor = vi.fn(),
	onCommit = vi.fn(),
) {
	act(() => {
		root.render(
			<RecentColorsProvider value={{ recentColors, pushColor }}>
				<RecentColorsRow {...props} onCommit={onCommit} />
			</RecentColorsProvider>,
		);
	});
	return { onCommit, pushColor };
}

describe('recentColorsRow (wave-4 B6)', () => {
	it('renders nothing while the list is empty', () => {
		mount([]);
		expect(row()).toBeNull();
	});

	it('is the shared grid with only Recent Colors, one named swatch per colour', () => {
		mount(['#112233', '#445566']);
		expect(row()?.localName).toBe('pptx-ui-color-grid');
		expect(row()?.getAttribute('aria-label')).toBe('Recent Colors');
		expect(
			[...row()!.shadowRoot!.querySelectorAll('.heading')].map((heading) => heading.textContent),
		).toStrictEqual(['Recent Colors']);
		expect(swatches().map((swatch) => swatch.getAttribute('aria-label'))).toStrictEqual([
			'Recent #112233',
			'Recent #445566',
		]);
		expect(row()!.shadowRoot!.querySelectorAll('button')).toHaveLength(2);
	});

	it('clicking a swatch commits the colour AND pushes it', () => {
		const { onCommit, pushColor } = mount(['#112233']);
		act(() => swatches()[0]!.click());
		expect(onCommit).toHaveBeenCalledWith('#112233');
		expect(pushColor).toHaveBeenCalledWith('#112233');
	});

	it('still commits after the list fills from empty, with the same callbacks', () => {
		const pushColor = vi.fn();
		const onCommit = vi.fn();
		mount([], {}, pushColor, onCommit);
		mount(['#112233'], {}, pushColor, onCommit);
		act(() => swatches()[0]!.click());
		expect(onCommit).toHaveBeenCalledWith('#112233');
		expect(pushColor).toHaveBeenCalledWith('#112233');
	});

	it('disables every swatch when disabled is set', () => {
		mount(['#112233'], { disabled: true });
		expect(swatches()[0]!.disabled).toBeTruthy();
	});
});

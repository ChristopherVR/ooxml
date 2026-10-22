// @vitest-environment happy-dom
/**
 * W3-G2: the theme colour picker grid renders the deck's REAL theme palette
 * (via `ooxml-ui/pptx`'s `buildThemeColorSwatchGrid`) instead of a
 * hard-coded Office palette, on the shared `pptx-ui-color-grid`, highlights
 * the swatch matching the element's current ref/hex, and commits both the
 * resolved hex and the ref on click.
 */
import type { ThemeColorPickerCommit } from 'ooxml-ui/pptx';
import React, { act } from 'react';
import { createRoot } from 'react-dom/client';
import type { Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it } from 'vitest';

import { ThemeColorMapProvider } from './ThemeColorMapContext';
import { ThemeColorSwatchGrid } from './ThemeColorSwatchGrid';

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

let container: HTMLDivElement;
let root: Root;

const grid = () => container.querySelector('pptx-ui-color-grid');
/** A Theme Colors swatch of the shared grid, by its label. */
const themeSwatch = (title: string) =>
	grid()?.shadowRoot?.querySelector<HTMLButtonElement>(`button[title="${title}"]`) ?? null;

beforeEach(() => {
	container = document.createElement('div');
	document.body.appendChild(container);
	root = createRoot(container);
});

afterEach(() => {
	act(() => root.unmount());
	container.remove();
});

describe('themeColorSwatchGrid', () => {
	it('renders nothing when no deck theme is loaded', () => {
		act(() => {
			root.render(
				<ThemeColorMapProvider value={undefined}>
					<ThemeColorSwatchGrid onPick={() => {}} />
				</ThemeColorMapProvider>,
			);
		});
		expect(grid()).toBeNull();
	});

	it('renders a base + variant swatch button per theme column', () => {
		act(() => {
			root.render(
				<ThemeColorMapProvider value={OFFICE_THEME}>
					<ThemeColorSwatchGrid onPick={() => {}} />
				</ThemeColorMapProvider>,
			);
		});
		const shadow = grid()!.shadowRoot!;
		// 10 columns x (1 base + 5 variants) = 60 swatch buttons, and no standard row.
		expect(shadow.querySelectorAll('button.swatch[data-source="theme"]')).toHaveLength(60);
		expect(shadow.querySelectorAll('button.swatch')).toHaveLength(60);
		// Only the Theme Colors heading: the inspector keeps its own standard and recent rows.
		expect(shadow.querySelectorAll('.heading')).toHaveLength(1);
	});

	it('highlights the swatch matching the element ref', () => {
		act(() => {
			root.render(
				<ThemeColorMapProvider value={OFFICE_THEME}>
					<ThemeColorSwatchGrid
						selectedRef={{ scheme: 'accent1', lumMod: 0.2, lumOff: 0.8 }}
						onPick={() => {}}
					/>
				</ThemeColorMapProvider>,
			);
		});
		const checked = grid()!.shadowRoot!.querySelectorAll('button[aria-checked="true"]');
		expect([...checked].map((el) => el.getAttribute('title'))).toStrictEqual([
			'Accent 1, Lighter 80%',
		]);
	});

	it('disables every swatch when the field is disabled', () => {
		act(() => {
			root.render(
				<ThemeColorMapProvider value={OFFICE_THEME}>
					<ThemeColorSwatchGrid disabled onPick={() => {}} />
				</ThemeColorMapProvider>,
			);
		});
		const swatches = [...grid()!.shadowRoot!.querySelectorAll('button.swatch')];
		expect(swatches.every((button) => (button as HTMLButtonElement).disabled)).toBeTruthy();
	});

	it('commits both the hex and the ref on click', () => {
		let commit: ThemeColorPickerCommit | undefined;
		act(() => {
			root.render(
				<ThemeColorMapProvider value={OFFICE_THEME}>
					<ThemeColorSwatchGrid onPick={(c) => (commit = c)} />
				</ThemeColorMapProvider>,
			);
		});
		const accent1Base = themeSwatch('Accent 1')!;
		expect(accent1Base).not.toBeNull();
		act(() => {
			accent1Base.click();
		});
		expect(commit).toStrictEqual({ hex: '#4472c4', ref: { scheme: 'accent1' } });
	});
});

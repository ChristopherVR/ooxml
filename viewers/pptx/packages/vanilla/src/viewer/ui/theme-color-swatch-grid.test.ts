/**
 * The inspector's Theme Colors grid is the shared `pptx-ui-color-grid`, fed the deck's real theme
 * palette: hidden until a theme is loaded, only the Theme Colors section (the callers keep their
 * own standard and recent rows), the stored ref or colour selected, and a pick committing both the
 * resolved hex and the ref.
 */
import type { ThemeColorPickerCommit } from 'ooxml-ui/pptx';
import { describe, expect, it, vi } from 'vitest';

import { createThemeColorSwatchGrid } from './theme-color-swatch-grid';
import { themeSwatch, themeSwatches } from './theme-color-swatch-grid.test-support';

const OFFICE_THEME: Record<string, string> = {
	dk1: '#000000',
	lt1: '#ffffff',
	dk2: '#44546a',
	lt2: '#e7e6e6',
	accent1: '#4472c4',
	accent2: '#ed7d31',
	accent3: '#a5a5a5',
	accent4: '#ffc000',
	accent5: '#5b9bd5',
	accent6: '#70ad47',
	bg1: '#ffffff',
	tx1: '#000000',
	bg2: '#e7e6e6',
	tx2: '#44546a',
};

const identity = (key: string): string => key;

function grid(onPick: (commit: ThemeColorPickerCommit) => void = vi.fn()) {
	const handle = createThemeColorSwatchGrid(document, identity, onPick);
	handle.setThemeColorMap(OFFICE_THEME);
	return handle;
}

const checkedTitles = (el: HTMLElement): (string | null)[] =>
	Array.from(el.shadowRoot!.querySelectorAll('button[aria-checked="true"]')).map((button) =>
		button.getAttribute('title'),
	);

describe('createThemeColorSwatchGrid', () => {
	it('is hidden until a theme is loaded, and again once it is cleared', () => {
		const handle = createThemeColorSwatchGrid(document, identity, vi.fn());
		expect(handle.el.hidden).toBeTruthy();

		handle.setThemeColorMap(OFFICE_THEME);
		expect(handle.el.hidden).toBeFalsy();

		handle.setThemeColorMap(undefined);
		expect(handle.el.hidden).toBeTruthy();
	});

	it('is the shared grid with a base and five variants per theme column, and no other section', () => {
		const { el } = grid();
		expect(el.localName).toBe('pptx-ui-color-grid');
		const shadow = el.shadowRoot!;
		// 10 columns x (1 base + 5 variants) = 60 swatches, and no standard row.
		expect(shadow.querySelectorAll('button.swatch[data-source="theme"]')).toHaveLength(60);
		expect(shadow.querySelectorAll('button.swatch')).toHaveLength(60);
		// Only the Theme Colors heading: the callers keep their own standard and recent rows.
		expect(
			Array.from(shadow.querySelectorAll('.heading')).map((heading) => heading.textContent),
		).toStrictEqual(['pptx.colorPicker.themeColors']);
		expect(el.getAttribute('aria-label')).toBe('pptx.colorPicker.themeColors');
	});

	it('draws each swatch in the colour the theme resolves', () => {
		const { el } = grid();
		const accent2 = themeSwatch(el, 'Accent 2')!;
		expect(accent2.dataset.color).toBe('#ed7d31');
		expect(accent2.getAttribute('aria-label')).toBe('Accent 2');
	});

	it('selects the swatch matching the stored ref, or else the colour', () => {
		const handle = grid();
		handle.setSelected({ scheme: 'accent1', lumMod: 0.2, lumOff: 0.8 }, undefined);
		expect(checkedTitles(handle.el)).toStrictEqual(['Accent 1, Lighter 80%']);

		handle.setSelected(undefined, '#ED7D31');
		expect(checkedTitles(handle.el)).toStrictEqual(['Accent 2']);

		handle.setSelected(undefined, undefined);
		expect(checkedTitles(handle.el)).toStrictEqual([]);
	});

	it('commits both the hex and the ref of a base or variant swatch', () => {
		const onPick = vi.fn();
		const { el } = grid(onPick);

		themeSwatch(el, 'Accent 1')!.click();
		expect(onPick).toHaveBeenLastCalledWith({ hex: '#4472c4', ref: { scheme: 'accent1' } });

		themeSwatch(el, 'Accent 1, Lighter 80%')!.click();
		expect(onPick).toHaveBeenCalledTimes(2);
		expect(onPick.mock.lastCall![0].ref).toStrictEqual({
			scheme: 'accent1',
			lumMod: 0.2,
			lumOff: 0.8,
		});
	});

	it('disables every swatch, and ignores clicks, while disabled', () => {
		const onPick = vi.fn();
		const handle = grid(onPick);

		handle.setDisabled(true);
		const swatches = themeSwatches(handle.el);
		expect(swatches).toHaveLength(60);
		expect(swatches.every((button) => button.disabled)).toBeTruthy();
		themeSwatch(handle.el, 'Accent 2')!.click();
		expect(onPick).not.toHaveBeenCalled();

		handle.setDisabled(false);
		expect(themeSwatches(handle.el).some((button) => button.disabled)).toBeFalsy();
		themeSwatch(handle.el, 'Accent 2')!.click();
		expect(onPick).toHaveBeenCalledExactlyOnceWith({ hex: '#ed7d31', ref: { scheme: 'accent2' } });
	});
});

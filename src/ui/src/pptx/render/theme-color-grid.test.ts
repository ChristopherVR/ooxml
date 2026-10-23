import { describe, expect, it } from 'vitest';

import { recentColorsGridView, themeColorGridCommit, themeColorGridView } from './theme-color-grid';
import { buildThemeColorSwatchGrid } from './theme-color-swatches';

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
const columns = buildThemeColorSwatchGrid(OFFICE_THEME);
const accent1Lighter80 = { scheme: 'accent1' as const, lumMod: 0.2, lumOff: 0.8 };

describe('theme colour grid view', () => {
	it('hands the grid the deck swatches as the slide resolves them', () => {
		const { palette } = themeColorGridView(columns, undefined, undefined);
		expect(palette).toHaveLength(10);
		expect(palette[4]!.base).toStrictEqual({ hex: '#4472c4', label: 'Accent 1' });
		const lighter80 = palette[4]!.variants[0]!;
		// The same hex as the canvas resolves for accent1 lumMod 0.2 lumOff 0.8.
		expect(lighter80.hex).toBe(columns[4]!.variants[0]!.hex.toLowerCase());
		expect(lighter80.label).toBe('Accent 1, Lighter 80%');
	});

	it('selects by the stored theme reference first, then by colour', () => {
		const byRef = themeColorGridView(columns, accent1Lighter80, '#000000');
		expect(byRef.selectedTheme).toStrictEqual({
			column: 4,
			variant: { kind: 'lighter', percent: 80 },
		});
		expect(themeColorGridView(columns, undefined, '#ED7D31').selectedTheme).toStrictEqual({
			column: 5,
		});
		expect(themeColorGridView(columns, undefined, '#123456').selectedTheme).toBeNull();
	});

	it('turns a theme pick back into the hex and reference to store', () => {
		const commit = themeColorGridCommit(columns, {
			color: '#dae3f3',
			source: 'theme',
			label: 'Accent 1, Lighter 80%',
			theme: { column: 4, variant: { kind: 'lighter', percent: 80 } },
		});
		expect(commit).toStrictEqual({ hex: columns[4]!.variants[0]!.hex, ref: accent1Lighter80 });
		expect(
			themeColorGridCommit(columns, {
				color: '#4472c4',
				source: 'theme',
				label: '',
				theme: { column: 4 },
			}),
		).toStrictEqual({ hex: columns[4]!.base.hex, ref: { scheme: 'accent1' } });
		expect(
			themeColorGridCommit(columns, { color: '#ff0000', source: 'standard', label: 'Red' }),
		).toBeUndefined();
	});

	it('shows nothing before a deck theme is loaded', () => {
		expect(themeColorGridView([], undefined, '#ffffff')).toStrictEqual({
			palette: [],
			selectedTheme: null,
		});
	});
});

describe('recentColorsGridView', () => {
	it('shows only Recent Colors, named in the viewer language', () => {
		const t = (key: string, params?: Record<string, string | number>) =>
			key === 'pptx.colorPicker.recentColors'
				? 'Zuletzt verwendete Farben'
				: `${key} ${params?.hex}`;
		const { recentLabel, ...view } = recentColorsGridView(['#FF0000', '#0000FF'], t);
		expect(view).toStrictEqual({
			palette: [],
			standardColors: [],
			recentColors: ['#FF0000', '#0000FF'],
			label: 'Zuletzt verwendete Farben',
			recentHeading: 'Zuletzt verwendete Farben',
		});
		expect(recentLabel('#FF0000')).toBe('pptx.colorPicker.recentColor #FF0000');
	});
});

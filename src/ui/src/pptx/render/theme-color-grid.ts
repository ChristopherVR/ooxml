import type { ThemePaletteColumn } from 'ooxml-core/color';
import type { PptxThemeColorRef } from 'ooxml-core/pptx';

import type { OfficeColorPick, OfficeThemeSwatchPosition } from '../../form/color-grid';
import {
	findSelectedThemeSwatch,
	themeSwatchCommit,
	type ThemeColorPickerCommit,
} from './theme-color-picker-state';
import type { ThemeColorSwatch, ThemeColorSwatchColumn } from './theme-color-swatches';

export type {
	OfficeColorPick,
	OfficeColorPickEvent,
	OfficeThemeSwatchPosition,
} from '../../form/color-grid';

/** What `pptx-ui-color-grid` shows for a deck: its own swatches and the current one. */
export interface ThemeColorGridView {
	palette: ThemePaletteColumn[];
	selectedTheme: OfficeThemeSwatchPosition | null;
}

type SteppedSwatch = ThemeColorSwatch & {
	readonly variant: NonNullable<ThemeColorSwatch['variant']>;
};
const stepped = (swatch: ThemeColorSwatch): swatch is SteppedSwatch => swatch.variant !== undefined;

// The grid draws five variants, each naming its step, under every column; a column the theme
// resolves always has them.
const drawable = (columns: readonly ThemeColorSwatchColumn[]) =>
	columns.filter((column) => column.variants.length === 5 && column.variants.every(stepped));

const paletteSwatch = (swatch: ThemeColorSwatch) => ({
	hex: swatch.hex.toLowerCase(),
	label: swatch.label,
});

/**
 * PowerPoint's Theme Colors for the shared grid: the swatches `buildThemeColorSwatchGrid` resolved
 * from the deck (so they match what the slide paints), and the swatch the stored theme reference,
 * or else the colour, selects.
 */
export function themeColorGridView(
	columns: readonly ThemeColorSwatchColumn[],
	selectedRef: PptxThemeColorRef | undefined,
	selectedHex: string | undefined,
): ThemeColorGridView {
	const shown = drawable(columns);
	const selected = findSelectedThemeSwatch(shown, selectedRef, selectedHex);
	let selectedTheme: OfficeThemeSwatchPosition | null = null;
	const palette = shown.map((column, index) => {
		if (selected === column.base) {
			selectedTheme = { column: index };
		}
		const variant = column.variants.find((swatch) => swatch === selected)?.variant;
		if (variant) {
			selectedTheme = { column: index, variant };
		}
		return {
			name: column.base.label,
			base: paletteSwatch(column.base),
			variants: column.variants
				.filter(stepped)
				.map((swatch) => ({ ...paletteSwatch(swatch), variant: swatch.variant })),
		};
	});
	return { palette, selectedTheme };
}

type Translate = (key: string, params?: Record<string, string | number>) => string;

/** What a recent-colours row's `pptx-ui-color-grid` shows: Recent Colors and nothing else. */
export interface RecentColorsGridView {
	palette: ThemePaletteColumn[];
	standardColors: [];
	recentColors: readonly string[];
	label: string;
	recentHeading: string;
	recentLabel: (hex: string) => string;
}

/**
 * The "Recent Colors" row under an inspector colour picker, as a `pptx-ui-color-grid` with only
 * that section: every binding sets these fields (and `data-testid="pptx-color-recent"`) and
 * renders nothing while the list is empty, so the row looks, reads and is named the same
 * everywhere. A pick arrives as `office-color-pick` with source `recent`.
 */
export function recentColorsGridView(
	colors: readonly string[],
	t: Translate,
): RecentColorsGridView {
	const heading = t('pptx.colorPicker.recentColors');
	return {
		palette: [],
		standardColors: [],
		recentColors: colors,
		label: heading,
		recentHeading: heading,
		recentLabel: (hex) => t('pptx.colorPicker.recentColor', { hex }),
	};
}

/** The hex and theme reference to store for a pick from that grid; undefined for any other pick. */
export function themeColorGridCommit(
	columns: readonly ThemeColorSwatchColumn[],
	pick: OfficeColorPick,
): ThemeColorPickerCommit | undefined {
	if (pick.source !== 'theme' || !pick.theme) {
		return undefined;
	}
	const { variant } = pick.theme;
	const column = drawable(columns)[pick.theme.column];
	const swatch = variant
		? column?.variants.find(
				(item) => item.variant?.kind === variant.kind && item.variant.percent === variant.percent,
			)
		: column?.base;
	return swatch ? themeSwatchCommit(swatch) : undefined;
}

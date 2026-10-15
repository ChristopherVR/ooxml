import { themeColor } from '../../chart/render/theme-palette';
import { applyTint } from '../layout/tint';
import type { ThemePalette } from '../model';

/** The built-in slicer styles, in the order Excel's Slicer Styles gallery lists them. */
export const SLICER_STYLE_NAMES: readonly string[] = [
	...[1, 2, 3, 4, 5, 6].map((n) => `SlicerStyleLight${n}`),
	...[1, 2, 3, 4, 5, 6].map((n) => `SlicerStyleDark${n}`),
	'SlicerStyleOther1',
	'SlicerStyleOther2',
];

export const DEFAULT_SLICER_STYLE = 'SlicerStyleLight1';

export interface SlicerStyleColors {
	background: string;
	border: string;
	headerBackground: string;
	headerText: string;
	headerBorder: string;
	selectedFill: string;
	selectedText: string;
	selectedBorder: string;
	unselectedFill: string;
	unselectedText: string;
	unselectedBorder: string;
	/** Items without data under the other slicers keep their state but fade. */
	noDataText: string;
	noDataSelectedFill: string;
}

const hex = (value: string) => `#${value}`;

/**
 * Colours of a built-in slicer style resolved against the workbook theme (accent 1 to 6 for the
 * Light and Dark families). The values follow the look of Excel's styles; they are not measured
 * from a native render. An unknown or custom style name falls back to Light 1.
 */
export function slicerStyleColors(
	name: string | undefined,
	theme: ThemePalette,
): SlicerStyleColors {
	const match = /^SlicerStyle(Light|Dark|Other)(\d)$/.exec(name ?? '') ?? ['', 'Light', '1'];
	const family = match[1];
	const index = Number(match[2]);
	const accentIndex = family === 'Other' ? (index === 1 ? 4 : 6) : 3 + index;
	const accent = themeColor(theme, accentIndex) ?? '5B9BD5';
	const white = themeColor(theme, 0) ?? 'FFFFFF';
	const dark = themeColor(theme, 1) ?? '000000';
	const gray = applyTint(white, -0.15);
	const tint = (tintValue: number) => hex(applyTint(accent, tintValue));
	if (family === 'Dark')
		return {
			background: hex(gray),
			border: hex(applyTint(gray, -0.25)),
			headerBackground: hex(accent),
			headerText: hex(white),
			headerBorder: hex(accent),
			selectedFill: hex(accent),
			selectedText: hex(white),
			selectedBorder: tint(-0.25),
			unselectedFill: hex(applyTint(white, -0.05)),
			unselectedText: hex(dark),
			unselectedBorder: hex(applyTint(white, -0.25)),
			noDataText: hex(applyTint(dark, 0.6)),
			noDataSelectedFill: tint(0.4),
		};
	return {
		background: hex(white),
		border: hex(applyTint(white, -0.35)),
		headerBackground: hex(white),
		headerText: hex(dark),
		headerBorder: tint(0),
		selectedFill: tint(0.6),
		selectedText: hex(dark),
		selectedBorder: tint(0),
		unselectedFill: hex(white),
		unselectedText: hex(dark),
		unselectedBorder: hex(applyTint(white, -0.25)),
		noDataText: hex(applyTint(dark, 0.6)),
		noDataSelectedFill: tint(0.8),
	};
}

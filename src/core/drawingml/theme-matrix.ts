import type { ThemeEffectStyle, ThemeFormatScheme } from './format-scheme';
import type { DrawingColor } from './types';

/** A self-contained style matrix for a theme with no authored fill or line styles. */
export function themeMatrix(
	name = 'Office',
	effectStyles: ThemeEffectStyle[] = Array.from({ length: 3 }, () => ({
		effects: { effects: [] },
	})),
): ThemeFormatScheme {
	const color: DrawingColor = { kind: 'scheme', value: 'phClr', transforms: [] };
	const fills = () =>
		Array.from({ length: 3 }, () => ({ fill: { kind: 'solid' as const, color } }));
	return {
		name,
		fillStyles: fills(),
		lineStyles: [6350, 12700, 19050].map((widthEmu) => ({
			widthEmu,
			fill: { kind: 'solid', color },
		})),
		effectStyles,
		backgroundFillStyles: fills(),
	};
}

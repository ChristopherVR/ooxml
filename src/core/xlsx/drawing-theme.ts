import { PALETTE_SLOTS } from './read/theme';
import type { ThemePalette } from './model';

/** The worksheet palette in DrawingML slot order, including the standard text/background aliases. */
export function drawingSchemeColors(theme: ThemePalette | undefined): Record<string, string> {
	const colors: Record<string, string> = {};
	PALETTE_SLOTS.forEach((slot, index) => {
		const value = theme?.colors[index];
		if (value && /^[0-9a-f]{6}$/i.test(value)) colors[slot] = `#${value.toLowerCase()}`;
	});
	for (const [alias, slot] of Object.entries({ bg1: 'lt1', tx1: 'dk1', bg2: 'lt2', tx2: 'dk2' })) {
		const value = colors[slot];
		if (value) colors[alias] = value;
	}
	return colors;
}

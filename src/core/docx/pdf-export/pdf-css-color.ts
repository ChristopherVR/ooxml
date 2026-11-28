// CSS colours as PDF fill and stroke channels, for the Word PDF export's painters. A module of
// its own so the WordArt painter, which loads on first use, shares it without the text writer.
import { parseColor, type Rgba } from '../../pdf/write/pdf-svg-style';

/** A CSS colour (or bare `RRGGBB`) as 0..1 channels; `null` for transparent or unknown. */
export function cssRgba(color: string | undefined): Rgba | null {
	if (!color) return null;
	try {
		return parseColor(/^[\da-f]{6}$/i.test(color) ? `#${color}` : color);
	} catch {
		return null;
	}
}

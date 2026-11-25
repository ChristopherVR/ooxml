import type { FontFallback } from '../font-table';
import { metricFontFamily } from '../../text/metric-font-family';
import { DEFAULT_FONT_FAMILY } from './units';

/** Word's pagination policy uses metric-compatible substitutes from shared text. */
export function cssFontStack(family: string | undefined, fallback?: FontFallback): string {
	return metricFontFamily(family?.trim() || DEFAULT_FONT_FAMILY, fallback);
}

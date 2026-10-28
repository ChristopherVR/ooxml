import { metricFontFamily } from '../../text/metric-font-family';
import { DEFAULT_FONT_FAMILY } from './units';

/** Word's pagination policy uses metric-compatible substitutes from shared text. */
export function cssFontStack(family: string | undefined): string {
	return metricFontFamily(family?.trim() || DEFAULT_FONT_FAMILY);
}

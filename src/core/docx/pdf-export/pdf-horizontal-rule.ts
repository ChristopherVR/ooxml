/** A Word horizontal line (`v:rect` with `o:hr`) for the PDF export: a filled rule, bevelled unless it is plain. DOM-free. */
import type { LayoutObject } from '../layout/index';
import type { PdfVectorPage } from '../../pdf/write/pdf-vector';
import { fillWith } from './pdf-text';

/** Where Word ends the rule above the bottom of its line: 2.5 pt (measured in Word 16's PDF). */
const RULE_BOTTOM_PX = 2.52 * (4 / 3);
/** Thickness of the bevel Word draws round a shaded rule, and its two greys (0.24 pt). */
const BEVEL_PX = 0.24 * (4 / 3);
const BEVEL_DARK = '#a0a0a0';
const BEVEL_LIGHT = '#e3e3e3';

/** Draws the rule of `object` in its box, whose top left is (x, y). */
export function paintHorizontalRule(
	page: PdfVectorPage,
	object: Pick<LayoutObject, 'widthPx' | 'heightPx' | 'rule'>,
	x: number,
	y: number,
): void {
	const { rule, widthPx: width } = object;
	if (!rule || !(width > 0)) return;
	const top = y + object.heightPx - RULE_BOTTOM_PX - rule.heightPx;
	if (fillWith(page, rule.color, '#a0a0a0')) {
		page.rect(x, top, width, rule.heightPx);
		page.op('f');
	}
	if (!rule.shaded || rule.heightPx < 2 * BEVEL_PX) return;
	const edges: [string, number, number, number, number][] = [
		[BEVEL_DARK, x, top, width - BEVEL_PX, BEVEL_PX],
		[BEVEL_DARK, x, top, BEVEL_PX, rule.heightPx - BEVEL_PX],
		[BEVEL_LIGHT, x, top + rule.heightPx - BEVEL_PX, width, BEVEL_PX],
		[BEVEL_LIGHT, x + width - BEVEL_PX, top, BEVEL_PX, rule.heightPx],
	];
	for (const [color, ex, ey, ew, eh] of edges) {
		fillWith(page, color);
		page.rect(ex, ey, ew, eh);
		page.op('f');
	}
}

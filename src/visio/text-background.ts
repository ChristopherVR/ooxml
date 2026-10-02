import type { VisioText } from './model.js';
import { number, type Cells, type Report } from './sheet.js';
import { color, type Resources } from './style.js';

/**
 * TextBkgnd uses one-based palette indices; 0 and 255 mean no background.
 * Literal saved RGB values already encode the resolved color, without an offset.
 * https://learn.microsoft.com/en-us/office/client-developer/visio/textbkgnd-cell-text-block-format-section
 */
export function textBackground(
	cells: Cells,
	resources: Resources,
	report: Report,
): Pick<VisioText, 'backgroundColor' | 'backgroundOpacity'> {
	const raw = cells.get('TextBkgnd')?.value;
	if (raw === undefined || raw === '0' || raw === '255') return {};
	const value = /^\d+$/.test(raw) ? String(Number(raw) - 1) : raw;
	const resolved = color(new Map([['TextBkgnd', { value }]]), 'TextBkgnd', '', resources, report);
	if (!resolved) return {};
	return {
		backgroundColor: resolved,
		backgroundOpacity: Math.max(0, Math.min(1, 1 - number(cells, 'TextBkgndTrans', 0, report))),
	};
}

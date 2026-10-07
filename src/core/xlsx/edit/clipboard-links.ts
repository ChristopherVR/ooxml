import { rangesIntersect, type CellAddress, type CellRange } from '../address.js';
import type { Hyperlink, Worksheet } from '../model.js';
import type { ClipboardCells } from './types.js';
import { subtractRange } from './range-math.js';
import { hyperlinkPolicy } from '../../opc/safe-href.js';

export function htmlHref(link: Hyperlink | undefined): string | undefined {
	const href = link?.target
		? `${link.target}${link.location ? `#${link.location}` : ''}`
		: link?.location
			? `#${link.location}`
			: undefined;
	return href && (href.startsWith('#') || hyperlinkPolicy(href).store) ? href : undefined;
}

export function hyperlinkFromHtmlAttributes(
	attrs: Map<string, string> | undefined,
	range: CellRange,
): Hyperlink | undefined {
	const href = attrs?.get('href');
	if (!href || (!href.startsWith('#') && !hyperlinkPolicy(href).store)) return undefined;
	const link: Hyperlink = { range };
	if (href.startsWith('#')) link.location = href.slice(1);
	else link.target = href;
	const tooltip = attrs?.get('title');
	if (tooltip) link.tooltip = tooltip;
	return link;
}

export function copyHyperlinks(sheet: Worksheet, area: CellRange): Hyperlink[] {
	return sheet.hyperlinks
		.filter((link) => rangesIntersect(link.range, area))
		.map((link) => ({
			...structuredClone(link),
			range: {
				start: {
					row: Math.max(area.start.row, link.range.start.row) - area.start.row,
					col: Math.max(area.start.col, link.range.start.col) - area.start.col,
				},
				end: {
					row: Math.min(area.end.row, link.range.end.row) - area.start.row,
					col: Math.min(area.end.col, link.range.end.col) - area.start.col,
				},
			},
		}));
}

/** Preserve the portion of a hyperlink span outside the area being replaced or moved. */
export function clearHyperlinks(sheet: Worksheet, area: CellRange): void {
	sheet.hyperlinks = sheet.hyperlinks.flatMap((link) =>
		subtractRange(link.range, area).map((range) => ({ ...link, range })),
	);
}

export function pasteHyperlinks(
	sheet: Worksheet,
	dest: CellRange,
	cells: ClipboardCells,
	transpose: boolean,
	skipBlanks: boolean,
	cutLocation?: (location: string) => string,
): void {
	if (!cells.hyperlinks) return;
	if (!skipBlanks) clearHyperlinks(sheet, dest);
	const point = (at: CellAddress) => ({
		row: dest.start.row + (transpose ? at.col : at.row),
		col: dest.start.col + (transpose ? at.row : at.col),
	});
	for (const link of cells.hyperlinks) {
		const range = { start: point(link.range.start), end: point(link.range.end) };
		if (skipBlanks) clearHyperlinks(sheet, range);
		const pasted = { ...structuredClone(link), range };
		if (pasted.location && cutLocation) pasted.location = cutLocation(pasted.location);
		sheet.hyperlinks.push(pasted);
	}
}

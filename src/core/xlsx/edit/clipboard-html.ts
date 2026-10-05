import type { CellRange } from '../address.js';
import type { CellStyle, Color, ThemePalette } from '../model.js';
import type { ClipboardCells } from './types.js';

export const escapeHtml = (text: string): string =>
	text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

/** A colour as CSS hex; theme colours ignore tint, legacy indexed colours are skipped. */
export function cssColor(color: Color | undefined, theme: ThemePalette): string | undefined {
	if (!color || color.auto) return undefined;
	if (color.rgb) return `#${color.rgb.slice(-6)}`;
	if (color.theme !== undefined) {
		const hex = theme.colors[color.theme];
		return hex ? `#${hex}` : undefined;
	}
	return undefined;
}

function cssFor(style: CellStyle | undefined, theme: ThemePalette): string {
	if (!style) return '';
	const css: string[] = [];
	const font = style.font;
	if (font.name) css.push(`font-family:${font.name}`);
	if (font.size) css.push(`font-size:${font.size}pt`);
	if (font.bold) css.push('font-weight:700');
	if (font.italic) css.push('font-style:italic');
	const decorations = [font.underline ? 'underline' : '', font.strike ? 'line-through' : ''].filter(
		Boolean,
	);
	if (decorations.length) css.push(`text-decoration:${decorations.join(' ')}`);
	const color = cssColor(font.color, theme);
	if (color && color !== '#000000') css.push(`color:${color}`);
	if (style.fill.type === 'pattern' && style.fill.pattern === 'solid') {
		const bg = cssColor(style.fill.fgColor, theme);
		if (bg) css.push(`background:${bg}`);
	}
	const h = style.alignment?.horizontal;
	if (h && h !== 'general') css.push(`text-align:${h === 'centerContinuous' ? 'center' : h}`);
	const v = style.alignment?.vertical;
	if (v) css.push(`vertical-align:${v === 'center' ? 'middle' : v}`);
	if (style.alignment?.wrapText) css.push('white-space:normal');
	if (style.numFmt !== 'General')
		css.push(`mso-number-format:"${style.numFmt.replace(/\\/g, '\\\\').replace(/"/g, '\\0022')}"`);
	return css.join(';');
}

/** An HTML table for the system clipboard, with inline styles and merged cells as spans. */
export function toHtml(cells: ClipboardCells, theme: ThemePalette): string {
	const covered = new Set<string>();
	const spans = new Map<string, CellRange>();
	for (const m of cells.merges) {
		spans.set(`${m.start.row},${m.start.col}`, m);
		for (let r = m.start.row; r <= m.end.row; r++)
			for (let c = m.start.col; c <= m.end.col; c++)
				if (r !== m.start.row || c !== m.start.col) covered.add(`${r},${c}`);
	}
	const rows: string[] = [];
	for (let r = 0; r < cells.rows; r++) {
		const tds: string[] = [];
		for (let c = 0; c < cells.cols; c++) {
			if (covered.has(`${r},${c}`)) continue;
			const cell = cells.data[r]?.[c] ?? null;
			const attrs: string[] = [];
			const span = spans.get(`${r},${c}`);
			if (span && span.end.row > span.start.row)
				attrs.push(`rowspan="${span.end.row - span.start.row + 1}"`);
			if (span && span.end.col > span.start.col)
				attrs.push(`colspan="${span.end.col - span.start.col + 1}"`);
			if (cell && typeof cell.value === 'number') attrs.push(`x:num="${cell.value}"`);
			const css = cssFor(cell?.style, theme);
			if (css) attrs.push(`style="${escapeHtml(css)}"`);
			const text = escapeHtml(cell?.text ?? '').replace(/\n/g, '<br>');
			tds.push(`<td${attrs.length ? ` ${attrs.join(' ')}` : ''}>${text}</td>`);
		}
		rows.push(`<tr>${tds.join('')}</tr>`);
	}
	return (
		'<html xmlns:x="urn:schemas-microsoft-com:office:excel"><head><meta charset="utf-8"></head><body>' +
		`<table border="0" cellpadding="0" cellspacing="0" style="border-collapse:collapse">${rows.join('')}</table>` +
		'</body></html>'
	);
}

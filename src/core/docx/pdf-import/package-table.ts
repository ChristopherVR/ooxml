/** Writes a rebuilt table (`tables.ts`) as `w:tbl`: grid, widths, borders, shading and merges. */
import { el } from '../load/legacy-doc/xml';
import type { Part, Writer } from './package';
import type { GridBorder } from './table-grid';
import type { TableBlock, TableCellBlock } from './tables';

const twips = (points: number) => Math.round(points * 20);

function border(side: string, value: GridBorder | undefined): string {
	if (!value) return el(`w:${side}`, { 'w:val': 'nil' });
	return el(`w:${side}`, {
		'w:val': 'single',
		// Eighths of a point, within what Word accepts for a line border.
		'w:sz': Math.min(96, Math.max(2, Math.round(value.width * 8))),
		'w:space': 0,
		'w:color': value.color.replace('#', '').toUpperCase(),
	});
}

function cellProperties(
	table: TableBlock,
	cell: TableCellBlock,
	merge: 'restart' | 'continue' | undefined,
	ruled: boolean,
) {
	const width = table.columns
		.slice(cell.col, cell.col + cell.colSpan)
		.reduce((sum, value) => sum + value, 0);
	const borders = ruled
		? `<w:tcBorders>${border('top', cell.borders.top)}${border('left', cell.borders.left)}${border('bottom', cell.borders.bottom)}${border('right', cell.borders.right)}</w:tcBorders>`
		: '';
	return (
		'<w:tcPr>' +
		el('w:tcW', { 'w:w': twips(width), 'w:type': 'dxa' }) +
		(cell.colSpan > 1 ? el('w:gridSpan', { 'w:val': cell.colSpan }) : '') +
		(merge === 'restart' ? el('w:vMerge', { 'w:val': 'restart' }) : merge ? '<w:vMerge/>' : '') +
		borders +
		(cell.fill
			? el('w:shd', {
					'w:val': 'clear',
					'w:color': 'auto',
					'w:fill': cell.fill.replace('#', '').toUpperCase(),
				})
			: '') +
		'</w:tcPr>'
	);
}

export function tableXml(writer: Writer, table: TableBlock, part: Part): string {
	const ruled = table.cells.some((cell) => Object.keys(cell.borders).length > 0);
	const total = table.columns.reduce((sum, value) => sum + value, 0);
	const padding = twips(table.padding);
	let xml =
		'<w:tbl><w:tblPr>' +
		el('w:tblW', { 'w:w': twips(total), 'w:type': 'dxa' }) +
		el('w:tblInd', { 'w:w': twips(table.indent ?? 0), 'w:type': 'dxa' }) +
		el('w:tblLayout', { 'w:type': 'fixed' }) +
		`<w:tblCellMar><w:left w:w="${padding}" w:type="dxa"/><w:right w:w="${padding}" w:type="dxa"/></w:tblCellMar>` +
		'</w:tblPr><w:tblGrid>' +
		table.columns.map((width) => el('w:gridCol', { 'w:w': twips(width) })).join('') +
		'</w:tblGrid>';
	table.rows.forEach((row, r) => {
		xml +=
			'<w:tr><w:trPr>' +
			el('w:trHeight', { 'w:val': twips(row.height), 'w:hRule': 'atLeast' }) +
			(row.header ? '<w:tblHeader/>' : '') +
			'</w:trPr>';
		let col = 0;
		while (col < table.columns.length) {
			const own = table.cells.find((cell) => cell.row === r && cell.col === col);
			const above = own
				? undefined
				: table.cells.find(
						(cell) => cell.col === col && cell.row < r && cell.row + cell.rowSpan > r,
					);
			const cell = own ?? above;
			if (!cell) {
				// A grid unit no cell covers (it should not happen): an empty cell keeps the grid.
				xml += `<w:tc><w:tcPr>${el('w:tcW', { 'w:w': twips(table.columns[col]!), 'w:type': 'dxa' })}</w:tcPr><w:p/></w:tc>`;
				col++;
				continue;
			}
			const merge = cell.rowSpan > 1 ? (own ? 'restart' : 'continue') : undefined;
			const body = own
				? cell.blocks.map((block) => writer.paragraph(block, part)).join('') || '<w:p/>'
				: '<w:p/>';
			xml += `<w:tc>${cellProperties(table, cell, merge, ruled)}${body}</w:tc>`;
			col += cell.colSpan;
		}
		xml += '</w:tr>';
	});
	return `${xml}</w:tbl>`;
}

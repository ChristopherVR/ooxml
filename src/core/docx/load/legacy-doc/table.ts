// Tables: the .doc grid, row and cell properties as `w:tbl`, including column and row spans,
// borders, shading, widths, row heights, header rows and nested tables.
import type {
	DocTableBlock,
	DocTableCellBlock,
	DocTableRowBlock,
} from '@christophervr/ole2/doc-model-types';
import { borderXml, shadingXml } from './borders';
import { el } from './xml';

/** Writes the blocks of a cell (paragraphs and nested tables). */
export type CellWriter = (cell: DocTableCellBlock) => string;

const VERTICAL = { top: undefined, center: 'center', bottom: 'bottom' } as const;

function tableBorders(first: DocTableRowBlock | undefined): string {
	const borders = first?.props?.borders;
	if (!borders) return '';
	const inner =
		borderXml('top', borders.top) +
		borderXml('left', borders.left) +
		borderXml('bottom', borders.bottom) +
		borderXml('right', borders.right) +
		borderXml('insideH', borders.insideH) +
		borderXml('insideV', borders.insideV);
	return inner ? `<w:tblBorders>${inner}</w:tblBorders>` : '';
}

function tableProperties(table: DocTableBlock): string {
	const first = table.rows[0];
	const grid = table.gridTwips;
	const width = (grid.at(-1) ?? 0) - (grid[0] ?? 0);
	const gap = first?.props?.gapHalfTwips;
	const alignment = first?.props?.alignment;
	let inner = el('w:tblW', { 'w:w': Math.max(0, width), 'w:type': 'dxa' });
	if (alignment && alignment !== 'left') inner += el('w:jc', { 'w:val': alignment });
	else if (grid[0]) inner += el('w:tblInd', { 'w:w': grid[0], 'w:type': 'dxa' });
	inner += tableBorders(first);
	inner += el('w:tblLayout', { 'w:type': 'fixed' });
	if (gap !== undefined)
		inner += `<w:tblCellMar>${el('w:left', { 'w:w': gap, 'w:type': 'dxa' })}${el('w:right', { 'w:w': gap, 'w:type': 'dxa' })}</w:tblCellMar>`;
	return `<w:tblPr>${inner}</w:tblPr>`;
}

function cellProperties(cell: DocTableCellBlock, grid: readonly number[]): string {
	const start = grid[cell.gridColumn] ?? 0;
	const end = grid[cell.gridColumn + cell.columnSpan] ?? start;
	const props = cell.props;
	let inner = el('w:tcW', { 'w:w': Math.max(0, end - start), 'w:type': 'dxa' });
	if (cell.columnSpan > 1) inner += el('w:gridSpan', { 'w:val': cell.columnSpan });
	if (cell.vMerge === 'restart') inner += el('w:vMerge', { 'w:val': 'restart' });
	else if (cell.vMerge === 'continue') inner += el('w:vMerge');
	if (props) {
		const borders =
			borderXml('top', props.borders.top) +
			borderXml('left', props.borders.left) +
			borderXml('bottom', props.borders.bottom) +
			borderXml('right', props.borders.right);
		if (borders) inner += `<w:tcBorders>${borders}</w:tcBorders>`;
		inner += shadingXml(props.shading);
		if (props.noWrap) inner += el('w:noWrap');
		const vertical = VERTICAL[props.verticalAlign];
		if (vertical) inner += el('w:vAlign', { 'w:val': vertical });
	}
	return `<w:tcPr>${inner}</w:tcPr>`;
}

function rowXml(row: DocTableRowBlock, grid: readonly number[], writeCell: CellWriter): string {
	const columns = grid.length - 1;
	const first = row.cells[0];
	const last = row.cells.at(-1);
	const before = first?.gridColumn ?? 0;
	const after = last ? Math.max(0, columns - (last.gridColumn + last.columnSpan)) : 0;
	let trPr = '';
	if (before > 0) trPr += el('w:gridBefore', { 'w:val': before });
	if (after > 0) trPr += el('w:gridAfter', { 'w:val': after });
	const props = row.props;
	if (props?.cantSplit) trPr += el('w:cantSplit');
	if (props && props.heightTwips !== 0)
		trPr += el('w:trHeight', {
			'w:val': Math.abs(props.heightTwips),
			'w:hRule': props.heightRule === 'auto' ? undefined : props.heightRule,
		});
	if (props?.isHeader) trPr += el('w:tblHeader');
	const cells = row.cells
		.map((cell) => {
			const content = cell.vMerge === 'continue' ? '<w:p/>' : writeCell(cell);
			return `<w:tc>${cellProperties(cell, grid)}${content}</w:tc>`;
		})
		.join('');
	return `<w:tr>${trPr ? `<w:trPr>${trPr}</w:trPr>` : ''}${cells}</w:tr>`;
}

/** A table block as `<w:tbl>`. */
export function tableXml(table: DocTableBlock, writeCell: CellWriter): string {
	const grid = table.gridTwips;
	const columns = grid
		.slice(1)
		.map((edge, index) => el('w:gridCol', { 'w:w': Math.max(0, edge - (grid[index] ?? 0)) }))
		.join('');
	const rows = table.rows.map((row) => rowXml(row, grid, writeCell)).join('');
	return `<w:tbl>${tableProperties(table)}<w:tblGrid>${columns}</w:tblGrid>${rows}</w:tbl>`;
}

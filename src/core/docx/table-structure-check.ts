// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Which tables the structure editors may restructure without losing XML they cannot rebuild.
import { children, first, isElement, named, type XmlElement } from './xml';

const only = (parent: XmlElement, allowed: string[]): boolean =>
	Array.from(parent.childNodes)
		.filter(isElement)
		.every((node) => allowed.some((name) => named(node, name)));

/** Paragraph content the editors can move, copy or drop whole: plain runs with no hidden inline XML. */
export function plainParagraphs(cell: XmlElement): boolean {
	for (const paragraph of children(cell, 'p')) {
		if (!only(paragraph, ['pPr', 'r'])) return false;
		for (const run of children(paragraph, 'r'))
			if (!only(run, ['rPr', 't', 'tab', 'br', 'cr', 'noBreakHyphen'])) return false;
	}
	return true;
}

/** Structural edits on rectangular tables: no merges, nested tables or row offsets. */
export function canEditTableStructure(table: XmlElement): boolean {
	if (!only(table, ['tblPr', 'tblGrid', 'tr'])) return false;
	const grid = first(table, 'tblGrid');
	if (grid && !only(grid, ['gridCol'])) return false;
	const rows = children(table, 'tr');
	const [firstRow] = rows;
	if (!firstRow) return false;
	const width = children(firstRow, 'tc').length;
	if (!width) return false;
	if (grid && children(grid, 'gridCol').length !== width) return false;
	for (const row of rows) {
		if (!only(row, ['trPr', 'tc']) || children(row, 'tc').length !== width) return false;
		const rowProps = first(row, 'trPr');
		if (rowProps && ['gridBefore', 'gridAfter', 'trPrChange'].some((key) => first(rowProps, key)))
			return false;
		for (const cell of children(row, 'tc')) {
			if (!only(cell, ['tcPr', 'p']) || !children(cell, 'p').length) return false;
			const props = first(cell, 'tcPr');
			if (props && ['gridSpan', 'vMerge', 'hMerge', 'tcPrChange'].some((key) => first(props, key)))
				return false;
			if (!plainParagraphs(cell)) return false;
		}
	}
	return true;
}

/**
 * Tables the merge-aware structure editors accept: any mix of `gridSpan`, `vMerge` and nested
 * tables, but no row offsets, tracked row or cell changes, `hMerge`, or content controls around
 * rows and cells. The caller still checks that the cells tile the grid (`tableGrid`).
 */
export function canRestructureTable(table: XmlElement): boolean {
	if (!only(table, ['tblPr', 'tblGrid', 'tr'])) return false;
	const grid = first(table, 'tblGrid');
	if (!grid || !only(grid, ['gridCol'])) return false;
	const rows = children(table, 'tr');
	if (!rows.length) return false;
	for (const row of rows) {
		if (!only(row, ['trPr', 'tc']) || !children(row, 'tc').length) return false;
		const rowProps = first(row, 'trPr');
		if (
			rowProps &&
			['gridBefore', 'gridAfter', 'wBefore', 'wAfter', 'trPrChange', 'ins', 'del'].some((key) =>
				first(rowProps, key),
			)
		)
			return false;
		for (const cell of children(row, 'tc')) {
			if (!only(cell, ['tcPr', 'p', 'tbl']) || !children(cell, 'p').length) return false;
			const props = first(cell, 'tcPr');
			if (
				props &&
				['hMerge', 'tcPrChange', 'cellIns', 'cellDel', 'cellMerge'].some((key) => first(props, key))
			)
				return false;
			if (!plainParagraphs(cell)) return false;
		}
	}
	return true;
}

import { type CellRange, normalizeRange, rangesIntersect } from '../address.js';
import { getCell } from '../cells.js';
import type { Table, TableColumn, Workbook } from '../model.js';
import { writeValue } from './cell-values.js';
import { type EditContext, displayText, sheetAt } from './context.js';
import { currentRegion } from './filter.js';
import { shiftFormulaInBand } from './band-formulas.js';
import { rewriteFormulas } from './shift-formulas.js';
import { shiftSheetContent } from './shift-sheet.js';

function nextTableName(workbook: Workbook): string {
	const taken = new Set(workbook.sheets.flatMap((s) => s.tables.map((t) => t.name.toLowerCase())));
	const names = new Set(workbook.definedNames.map((n) => n.name.toLowerCase()));
	for (let n = 1; ; n++) if (!taken.has(`table${n}`) && !names.has(`table${n}`)) return `Table${n}`;
}

/** Unique header names: blanks become `ColumnN`, repeats get a number suffix (`Name2`). */
export function uniqueHeaders(texts: string[]): string[] {
	const used = new Set<string>();
	return texts.map((text, i) => {
		let base = text.trim() || `Column${i + 1}`;
		if (!used.has(base.toLowerCase())) {
			used.add(base.toLowerCase());
			return base;
		}
		for (let n = 2; ; n++) {
			const candidate = `${base}${n}`;
			if (!used.has(candidate.toLowerCase())) {
				used.add(candidate.toLowerCase());
				base = candidate;
				return base;
			}
		}
	});
}

/**
 * Turns a range into a table (a single cell expands to its current region). Without a header
 * row, cells are shifted down one row and a `Column1`, `Column2`, ... header is inserted.
 */
export function createTable(
	ctx: EditContext,
	s: number,
	range: CellRange,
	hasHeader: boolean,
	styleName = 'TableStyleMedium2',
): Table {
	const { workbook } = ctx;
	const sheet = sheetAt(workbook, s);
	let r = normalizeRange(range);
	if (r.start.row === r.end.row && r.start.col === r.end.col) r = currentRegion(sheet, r.start);
	if (sheet.tables.some((t) => rangesIntersect(t.range, r)))
		throw new Error('A table cannot overlap another table.');
	if (sheet.merges.some((m) => rangesIntersect(m, r)))
		throw new Error('A table cannot contain merged cells.');
	return ctx.run(
		'Create table',
		'structure',
		[hasHeader ? { kind: 'sheet', sheet: s } : { kind: 'workbook' }],
		() => {
			if (!hasHeader) {
				const shift = { axis: 'row' as const, at: r.start.row, count: 1 };
				const band = { lo: r.start.col, hi: r.end.col };
				rewriteFormulas(workbook, (f, fs) => shiftFormulaInBand(f, fs, sheet.name, shift, band));
				shiftSheetContent(sheet, shift, band);
				r = { start: r.start, end: { row: r.end.row + 1, col: r.end.col } };
			}
			const width = r.end.col - r.start.col + 1;
			const texts = Array.from({ length: width }, (_v, i) =>
				hasHeader ? displayText(workbook, getCell(sheet, r.start.row, r.start.col + i)) : '',
			);
			const headers = uniqueHeaders(texts);
			const columns: TableColumn[] = headers.map((name, i) => {
				const cell = getCell(sheet, r.start.row, r.start.col + i);
				if (cell?.value !== name) {
					const styleId = cell?.styleId;
					writeValue(sheet, r.start.row, r.start.col + i, name);
					const written = getCell(sheet, r.start.row, r.start.col + i);
					if (written && styleId) written.styleId = styleId;
				}
				return { name };
			});
			const name = nextTableName(workbook);
			const id = Math.max(0, ...workbook.sheets.flatMap((ws) => ws.tables.map((t) => t.id))) + 1;
			const table: Table = {
				id,
				name,
				displayName: name,
				range: r,
				headerRow: true,
				totalsRow: false,
				columns,
				styleName,
				showRowStripes: true,
				showColumnStripes: false,
				showFirstColumn: false,
				showLastColumn: false,
			};
			sheet.tables.push(table);
			return table;
		},
		{ sheet: s, ranges: [r], structural: !hasHeader },
	);
}

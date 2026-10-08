import type { Workbook, Worksheet } from '../model';
import type { CellAddress } from '../address';
import { getCell } from '../cells';
import type { CellPosition } from './deps';
import { rewritePreservedFormulas } from './shift-preserved';
import { rewriteConditionalRule } from './conditional-formulas';

/**
 * Rewrites one formula; `formulaSheet` is the sheet unqualified references point at.
 * `at` identifies the origin of a cell-based formula, including its enclosing table.
 */
export type FormulaRewrite = (formula: string, formulaSheet: string, at?: CellAddress) => string;

/** Applies `rewrite`, keeping the original when the formula cannot be parsed. */
function safe(
	rewrite: FormulaRewrite,
	formula: string,
	formulaSheet: string,
	at?: CellAddress,
): string {
	try {
		return rewrite(formula, formulaSheet, at);
	} catch {
		return formula;
	}
}

/**
 * Visits every formula the workbook keeps: cell formulas, conditional-format and validation
 * formulas, defined names, chart series references and in-workbook hyperlink locations.
 * `formulaCells` can reuse a history capture, provided it also includes newly written formulas.
 */
export function rewriteFormulas(
	workbook: Workbook,
	rewrite: FormulaRewrite,
	formulaCells?: readonly CellPosition[],
): void {
	const bySheet = formulaCells && new Map<number, CellAddress[]>();
	if (bySheet)
		for (const cell of formulaCells ?? []) {
			const positions = bySheet.get(cell.sheet);
			if (positions) positions.push(cell);
			else bySheet.set(cell.sheet, [cell]);
		}
	workbook.sheets.forEach((sheet, index) =>
		rewriteSheetFormulas(sheet, rewrite, bySheet ? (bySheet.get(index) ?? []) : undefined),
	);
	for (const definedName of workbook.definedNames) {
		const local =
			definedName.localSheet === undefined ? undefined : workbook.sheets[definedName.localSheet];
		definedName.formula = safe(rewrite, definedName.formula, local?.name ?? '');
	}
}

/**
 * The per-sheet part of {@link rewriteFormulas}: cell, conditional-format, validation, table
 * column (calculated and totals) formulas, hyperlink locations, chart series references and the
 * `xm:f` formulas of preserved extensions (sparkline sources, x14 formats and validations).
 */
export function rewriteSheetFormulas(
	sheet: Worksheet,
	rewrite: FormulaRewrite,
	formulaCells?: readonly CellAddress[],
): void {
	rewritePreservedFormulas(sheet, rewrite);
	{
		const name = sheet.name;
		for (const table of sheet.tables)
			for (const [index, column] of table.columns.entries())
				if (column.calculatedColumnFormula !== undefined)
					column.calculatedColumnFormula = safe(rewrite, column.calculatedColumnFormula, name, {
						row: table.range.start.row + (table.headerRow ? 1 : 0),
						col: table.range.start.col + index,
					});
		if (formulaCells) {
			for (const at of formulaCells) {
				const cell = getCell(sheet, at.row, at.col);
				if (cell?.formula !== undefined) cell.formula = safe(rewrite, cell.formula, name, at);
			}
		} else {
			// Map callbacks avoid an entry-array allocation for every cell on large sheets.
			sheet.rows.forEach((cells, row) =>
				cells.forEach((cell, col) => {
					if (cell.formula !== undefined)
						cell.formula = safe(rewrite, cell.formula, name, { row, col });
				}),
			);
		}
		for (const cf of sheet.conditionalFormats)
			for (const rule of cf.rules)
				rewriteConditionalRule(rule, (f) => rewrite(f, name, cf.ranges[0]?.start));
		for (const dv of sheet.dataValidations) {
			if (dv.formula1 !== undefined)
				dv.formula1 = safe(rewrite, dv.formula1, name, dv.ranges[0]?.start);
			if (dv.formula2 !== undefined)
				dv.formula2 = safe(rewrite, dv.formula2, name, dv.ranges[0]?.start);
		}
		for (const link of sheet.hyperlinks)
			if (link.location && /![A-Za-z$]/.test(link.location))
				link.location = safe(rewrite, link.location, name);
		for (const drawing of sheet.drawings) {
			if (drawing.kind !== 'chart') continue;
			for (const series of drawing.series) {
				if (series.nameRef) series.nameRef = safe(rewrite, series.nameRef, name);
				if (series.categoriesRef) series.categoriesRef = safe(rewrite, series.categoriesRef, name);
				if (series.valuesRef) series.valuesRef = safe(rewrite, series.valuesRef, name);
			}
		}
	}
}

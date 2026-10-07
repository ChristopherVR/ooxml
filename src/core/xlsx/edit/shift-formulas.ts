import type { Workbook, Worksheet } from '../model.js';
import { rewritePreservedFormulas } from './shift-preserved.js';
import { rewriteConditionalRule } from './conditional-formulas.js';

/** Rewrites one formula; `formulaSheet` is the sheet unqualified references point at. */
export type FormulaRewrite = (formula: string, formulaSheet: string) => string;

/** Applies `rewrite`, keeping the original when the formula cannot be parsed. */
function safe(rewrite: FormulaRewrite, formula: string, formulaSheet: string): string {
	try {
		return rewrite(formula, formulaSheet);
	} catch {
		return formula;
	}
}

/**
 * Visits every formula the workbook keeps: cell formulas, conditional-format and validation
 * formulas, defined names, chart series references and in-workbook hyperlink locations.
 */
export function rewriteFormulas(workbook: Workbook, rewrite: FormulaRewrite): void {
	for (const sheet of workbook.sheets) rewriteSheetFormulas(sheet, rewrite);
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
export function rewriteSheetFormulas(sheet: Worksheet, rewrite: FormulaRewrite): void {
	rewritePreservedFormulas(sheet, rewrite);
	{
		const name = sheet.name;
		for (const table of sheet.tables)
			for (const column of table.columns)
				if (column.calculatedColumnFormula !== undefined)
					column.calculatedColumnFormula = safe(rewrite, column.calculatedColumnFormula, name);
		for (const cells of sheet.rows.values())
			for (const cell of cells.values())
				if (cell.formula !== undefined) cell.formula = safe(rewrite, cell.formula, name);
		for (const cf of sheet.conditionalFormats)
			for (const rule of cf.rules) rewriteConditionalRule(rule, (f) => rewrite(f, name));
		for (const dv of sheet.dataValidations) {
			if (dv.formula1 !== undefined) dv.formula1 = safe(rewrite, dv.formula1, name);
			if (dv.formula2 !== undefined) dv.formula2 = safe(rewrite, dv.formula2, name);
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

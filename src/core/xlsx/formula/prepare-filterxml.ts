import { prepareXPath } from '#xpath-runtime';
import type { Workbook } from '../model';
import { normalizeFunctionName } from './ast';
import { tokenize } from './tokenizer';

function usesFilterXml(formula: string | undefined): boolean {
	if (!formula || !/filterxml/i.test(formula)) return false;
	try {
		return tokenize(formula).some(
			(token) =>
				(token.kind === 'func' || token.kind === 'name') &&
				normalizeFunctionName(String(token.value)) === 'FILTERXML',
		);
	} catch {
		return false;
	}
}

/** Prepare an ad hoc formula before evaluating it synchronously in a browser. */
export function prepareFormula(formula: string): Promise<void> | undefined {
	if (usesFilterXml(formula)) return prepareXPath();
}

/** Check the formulas once at workbook open, before any synchronous calculation. */
export function prepareWorkbookFormulas(workbook: Workbook): Promise<void> | undefined {
	if (workbook.definedNames.some(({ formula }) => usesFilterXml(formula)))
		return prepareXPath(workbook);
	for (const sheet of workbook.sheets)
		for (const row of sheet.rows.values())
			for (const { formula } of row.values())
				if (usesFilterXml(formula)) return prepareXPath(workbook);
}

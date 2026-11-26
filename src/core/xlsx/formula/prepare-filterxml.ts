import { prepareXPath } from '#xpath-runtime';
import type { CfvoThreshold, ConditionalRule, Workbook } from '../model';
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

const thresholdUsesFilterXml = (threshold: CfvoThreshold): boolean =>
	threshold.type === 'formula' && usesFilterXml(threshold.value);

function ruleUsesFilterXml(rule: ConditionalRule): boolean {
	switch (rule.type) {
		case 'expression':
			return usesFilterXml(rule.formula);
		case 'cellIs':
			return rule.formulas.some(usesFilterXml);
		case 'colorScale':
		case 'iconSet':
			return rule.thresholds.some(thresholdUsesFilterXml);
		case 'dataBar':
			return thresholdUsesFilterXml(rule.min) || thresholdUsesFilterXml(rule.max);
		default:
			return false;
	}
}

/** Check the formulas once at workbook open, before any synchronous calculation. */
export function prepareWorkbookFormulas(workbook: Workbook): Promise<void> | undefined {
	if (workbook.definedNames.some(({ formula }) => usesFilterXml(formula)))
		return prepareXPath(workbook);
	for (const sheet of workbook.sheets) {
		if (
			sheet.conditionalFormats.some(({ rules }) => rules.some(ruleUsesFilterXml)) ||
			sheet.dataValidations.some(
				({ formula1, formula2 }) => usesFilterXml(formula1) || usesFilterXml(formula2),
			)
		)
			return prepareXPath(workbook);
		for (const row of sheet.rows.values())
			for (const { formula } of row.values())
				if (usesFilterXml(formula)) return prepareXPath(workbook);
	}
}

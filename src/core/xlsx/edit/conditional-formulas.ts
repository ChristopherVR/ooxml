import type { ConditionalRule } from '../model';
import { rewriteConditionalExtension } from '../conditional-extensions';

/** Visits modeled rule formulas, including formula-valued visual thresholds. */
export function rewriteConditionalRule(
	rule: ConditionalRule,
	rewrite: (formula: string) => string,
): void {
	const safe = (formula: string) => {
		try {
			return rewrite(formula);
		} catch {
			return formula;
		}
	};
	if (rule.type === 'cellIs') rule.formulas = rule.formulas.map(safe);
	else if (rule.type === 'expression') rule.formula = safe(rule.formula);
	if (rule.type === 'colorScale' || rule.type === 'iconSet')
		for (const threshold of rule.thresholds)
			if (threshold.type === 'formula' && threshold.value) threshold.value = safe(threshold.value);
	if (rule.type === 'dataBar')
		for (const threshold of [rule.min, rule.max])
			if (threshold.type === 'formula' && threshold.value) threshold.value = safe(threshold.value);
	if (rule.type === 'dataBar') rewriteConditionalExtension(rule, safe);
}

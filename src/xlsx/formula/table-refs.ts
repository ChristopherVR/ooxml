// Rewriting structured (table) references when a table is renamed.
import { FormulaError } from './ast.js';
import { joinTokens, type Token, tokenize } from './tokenizer.js';

function safeTokens(formula: string): Token[] | undefined {
	try {
		return tokenize(formula);
	} catch (e) {
		if (e instanceof FormulaError) return undefined;
		throw e;
	}
}

/**
 * Renames table `oldName` (case-insensitive) to `newName` in structured references
 * (`Old[Col]`, `Old[[#Totals],[Col]]`) and bare table names (`=SUM(Old)`). Sheet-qualified
 * names and text inside strings are left alone. Unparsable formulas are returned unchanged.
 */
export function renameTableInFormula(formula: string, oldName: string, newName: string): string {
	const tokens = safeTokens(formula);
	if (!tokens) return formula;
	const lower = oldName.toLowerCase();
	let changed = false;
	for (const token of tokens) {
		if (typeof token.value !== 'string' || token.value.toLowerCase() !== lower) continue;
		if (token.kind === 'structured' && token.text.startsWith(token.value)) {
			token.text = newName + token.text.slice(token.value.length);
			changed = true;
		} else if (token.kind === 'name' && !token.prefix && token.text === token.value) {
			token.text = newName;
			changed = true;
		}
	}
	return changed ? joinTokens(tokens) : formula;
}

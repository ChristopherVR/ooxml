// Rewriting structured (table) references when a table is renamed.
import { FormulaError } from './ast';
import { joinTokens, type Token, tokenize } from './tokenizer';

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

/** Renames a column, including unqualified references when the formula is inside its table. */
export function renameTableColumnInFormula(
	formula: string,
	tableName: string,
	oldName: string,
	newName: string,
	insideTable = false,
): string {
	const tokens = safeTokens(formula);
	if (!tokens) return formula;
	const escaped = newName.replace(/['#@[\]]/g, "'$&");
	let changed = false;
	for (const [index, token] of tokens.entries()) {
		if (token.kind !== 'structured' || token.prefix) continue;
		// External table references are tokenized as a prefixed name followed by bare brackets.
		if (token.value === '' && tokens[index - 1]?.prefix) continue;
		if (
			token.value === ''
				? !insideTable
				: String(token.value).toLowerCase() !== tableName.toLowerCase()
		)
			continue;
		// Match the innermost column brackets, preserving item selectors and range punctuation.
		const next = token.text.replace(
			/\[(@?)((?:[^\[\]']|'[\s\S])*)\]/g,
			(original, at: string, name: string) => {
				if (
					name.startsWith('#') ||
					name.replace(/'(.)/g, '$1').toLowerCase() !== oldName.toLowerCase()
				)
					return original;
				return `[${at}${escaped}]`;
			},
		);
		if (next !== token.text) {
			token.text = next;
			changed = true;
		}
	}
	return changed ? joinTokens(tokens) : formula;
}

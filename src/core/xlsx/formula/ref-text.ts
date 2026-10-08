// Reference text: tokenizing without throwing, and formatting rewritten references and sheet
// prefixes the way the formula rewriters (shift, rename, move, translate) write them.
import { columnLabel } from '../address';
import { FormulaError, type RefCorner, type RefSpec, type SheetPrefix } from './ast';
import { type Token, tokenize } from './tokenizer';

export function safeTokens(formula: string): Token[] | undefined {
	try {
		return tokenize(formula);
	} catch (e) {
		if (e instanceof FormulaError) return undefined;
		throw e;
	}
}

const cornerText = (c: RefCorner, kind: RefSpec['kind']): string => {
	if (kind === 'cols') return `${c.colAbs ? '$' : ''}${columnLabel(c.col)}`;
	if (kind === 'rows') return `${c.rowAbs ? '$' : ''}${c.row + 1}`;
	return `${c.colAbs ? '$' : ''}${columnLabel(c.col)}${c.rowAbs ? '$' : ''}${c.row + 1}`;
};

/** Formats a reference spec (without sheet prefix). */
export function formatRefSpec(spec: RefSpec): string {
	const a = cornerText(spec.start, spec.kind);
	if (spec.kind === 'cell') return a;
	return `${a}:${cornerText(spec.end, spec.kind)}`;
}

/** The text of a reference token pointing at `ref` instead (`#REF!` when undefined). */
export function rewriteRef(
	token: { prefix?: SheetPrefix; spill?: boolean },
	ref: RefSpec | undefined,
): string {
	const prefix = token.prefix?.text ?? '';
	if (!ref) return `${prefix}#REF!`;
	return `${prefix}${formatRefSpec(ref)}${token.spill ? '#' : ''}`;
}

export const sameSpec = (a: RefSpec, b: RefSpec): boolean =>
	a.kind === b.kind &&
	a.start.row === b.start.row &&
	a.start.col === b.start.col &&
	a.end.row === b.end.row &&
	a.end.col === b.end.col;

/**
 * Whether a sheet name must be quoted in a reference: anything but a plain identifier, and names
 * Excel would read as something else: A1 or R1C1 references (`R`, `C`, `RC`, `R1`, `C2`,
 * `R1C1`) and the logicals TRUE and FALSE.
 */
const needsQuote = (name: string): boolean =>
	!/^[A-Za-z_¡-￿][A-Za-z0-9_.¡-￿]*$/.test(name) ||
	/^[A-Za-z]{1,3}\d+$/.test(name) ||
	/^(?:R\d*C?\d*|C\d*)$/i.test(name) ||
	/^(?:TRUE|FALSE)$/i.test(name);

/** The text of a sheet prefix (`Sheet1!`, `'My Sheet'!`), quoted when needed. */
export function prefixText(prefix: SheetPrefix): string {
	const sheets = prefix.sheet2 !== undefined ? `${prefix.sheet}:${prefix.sheet2}` : prefix.sheet;
	const book = prefix.book !== undefined ? `[${prefix.book}]` : '';
	const quote =
		needsQuote(prefix.sheet) || (prefix.sheet2 !== undefined && needsQuote(prefix.sheet2));
	return quote ? `'${(book + sheets).replace(/'/g, "''")}'!` : `${book}${sheets}!`;
}

// Optional data validation for Replace and Replace All. Excel's own Replace never applies data
// validation, so hosts and the Find and Replace dialog opt in; this decides which replacements pass.
import {
	formatAddress,
	parseCellInput,
	replaceText,
	type FindMatch,
	type FindQuery,
} from '../index';
import type { EditSession } from '../edit/types';

/** A replacement that data validation refused, so the cell kept its value. */
export interface RefusedReplacement {
	sheet: number;
	row: number;
	col: number;
	/** The cell in A1 notation. */
	ref: string;
	/** What the cell would have held (formula-bar text). */
	value: string;
	/** The validation rule's error message. */
	message: string;
	title?: string;
}

/** Options of Replace and Replace All beyond the query and the replacement. */
export interface ReplaceOptions {
	/**
	 * Check each replaced value against the cell's data validation and leave cells it refuses
	 * unchanged. Default: the editor's `validateReplace` (false, as in Excel).
	 */
	validate?: boolean;
	/** Called once per call with the cells validation refused (not called when there are none). */
	onRefused?: (refused: RefusedReplacement[]) => void;
}

export interface ValidatedMatches {
	allowed: FindMatch[];
	refused: RefusedReplacement[];
}

/**
 * Splits matches into those whose replaced value passes the cell's data validation and those it
 * refuses. `matches` must come from a query that looks in formulas (the text a replace edits). A
 * formula result, a comment edit or a cell without a validation rule passes.
 */
export function validateReplacements(
	session: Pick<EditSession, 'validate' | 'workbook'>,
	query: FindQuery,
	replacement: string,
	matches: readonly FindMatch[],
): ValidatedMatches {
	const result: ValidatedMatches = { allowed: [], refused: [] };
	for (const match of matches) {
		const next = replaceText(match.text, query, replacement);
		if (query.lookIn === 'comments' || next.startsWith('=') || next === match.text) {
			result.allowed.push(match);
			continue;
		}
		const parsed = parseCellInput(next, { date1904: session.workbook.date1904 });
		const check = session.validate(match.sheet, match.row, match.col, parsed.value);
		if (check.ok) result.allowed.push(match);
		else
			result.refused.push({
				sheet: match.sheet,
				row: match.row,
				col: match.col,
				ref: formatAddress({ row: match.row, col: match.col }),
				value: next,
				message: check.message,
				...(check.title ? { title: check.title } : {}),
			});
	}
	return result;
}

/** `session.replaceAll` that skips the cells data validation refuses; reports them. */
export function replaceAllValidated(
	session: EditSession,
	query: FindQuery,
	replacement: string,
): { replaced: number; refused: RefusedReplacement[] } {
	const { allowed, refused } = validateReplacements(
		session,
		query,
		replacement,
		session.findAll(query),
	);
	if (!refused.length) return { replaced: session.replaceAll(query, replacement), refused };
	let replaced = 0;
	session.batch('Replace All', () => {
		for (const match of allowed) if (session.replaceOne(query, replacement, match)) replaced++;
	});
	return { replaced, refused };
}

/** `session.replaceOne` unless data validation refuses the new value; reports the refusal. */
export function replaceOneValidated(
	session: EditSession,
	query: FindQuery,
	replacement: string,
	at: FindMatch,
): { replaced: boolean; refused: RefusedReplacement[] } {
	const { allowed, refused } = validateReplacements(session, query, replacement, [at]);
	return { replaced: allowed.length > 0 && session.replaceOne(query, replacement, at), refused };
}

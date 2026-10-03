// Equality lookups for criteria over ranges many formulas share: SUMIF(A:A,A2,B:B) filled down a
// column asks the same cached block for a different value each time, so the block is indexed once
// by value instead of being scanned by every formula.
import { round15, parseNumberText } from '../text-number.js';
import type { Matrix, Scalar } from '../values.js';
import { hasWildcards } from './helpers.js';

const uses = new WeakMap<Matrix, number>();
const indexes = new WeakMap<Matrix, Map<string, number[]>>();

const numberKey = (n: number): string => `n${round15(n)}`;

/**
 * The key of an equality criterion (`7`, `"apple"`, `"=12"`), or `undefined` for criteria that
 * are not a plain match (comparisons, wildcards, blanks, logicals, errors).
 */
export function equalityKey(criteria: Scalar): string | undefined {
	if (typeof criteria === 'number') return numberKey(criteria);
	if (typeof criteria !== 'string') return undefined;
	const text = criteria.startsWith('=') ? criteria.slice(1) : criteria;
	if (/^(<|>)/.test(text) || text === '' || hasWildcards(text) || text.startsWith('#'))
		return undefined;
	const upper = text.toUpperCase();
	if (upper === 'TRUE' || upper === 'FALSE') return undefined;
	const n = parseNumberText(text);
	return n !== undefined ? numberKey(n) : `s${text.toLowerCase()}`;
}

/** The keys a cell value matches: numbers by value, text by lower case and as a number. */
function valueKeys(value: Scalar, out: string[]): void {
	out.length = 0;
	if (typeof value === 'number') out.push(numberKey(value));
	else if (typeof value === 'string') {
		out.push(`s${value.toLowerCase()}`);
		const n = parseNumberText(value);
		if (n !== undefined) out.push(numberKey(n));
	}
}

/**
 * Row-major positions in `block` whose value matches `key`, once the block has been searched
 * before (a block read once is cheaper to scan than to index). `undefined` means "scan".
 */
export function positionsOf(block: Matrix, key: string): number[] | undefined {
	let index = indexes.get(block);
	if (!index) {
		const count = (uses.get(block) ?? 0) + 1;
		uses.set(block, count);
		if (count < 2 || block.isPadded) return undefined;
		index = new Map();
		const keys: string[] = [];
		const cols = block.cols;
		for (let r = 0; r < block.rows; r++) {
			for (let c = 0; c < cols; c++) {
				valueKeys(block.get(r, c), keys);
				for (const k of keys) {
					const list = index.get(k);
					if (list) list.push(r * cols + c);
					else index.set(k, [r * cols + c]);
				}
			}
		}
		indexes.set(block, index);
	}
	return index.get(key) ?? [];
}

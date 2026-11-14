import { describe, expect, it } from 'vitest';
import { diffKeys, type DiffOp } from './diff';
import { slideRegions } from './slide';
import type { Token } from './tokens';

/** Longest common subsequence length by dynamic programming, the reference for small inputs. */
function lcs(a: string[], b: string[]): number {
	const table = Array.from({ length: a.length + 1 }, () => new Array<number>(b.length + 1).fill(0));
	for (let i = 1; i <= a.length; i++)
		for (let j = 1; j <= b.length; j++)
			table[i]![j] =
				a[i - 1] === b[j - 1]
					? table[i - 1]![j - 1]! + 1
					: Math.max(table[i - 1]![j]!, table[i]![j - 1]!);
	return table[a.length]![b.length]!;
}

function check(a: string[], b: string[], ops: DiffOp[]): void {
	let i = 0;
	let j = 0;
	for (const op of ops) {
		if (op.kind === 'equal') {
			expect([op.aIndex, op.bIndex]).toEqual([i++, j++]);
			expect(a[op.aIndex]).toBe(b[op.bIndex]);
		} else if (op.kind === 'delete') expect(op.aIndex).toBe(i++);
		else expect(op.bIndex).toBe(j++);
	}
	expect([i, j]).toEqual([a.length, b.length]);
}

describe('diffKeys', () => {
	it('finds a shortest edit script, deletions before insertions', () => {
		const ops = diffKeys(['a', 'b', 'c'], ['a', 'x', 'c']);
		expect(ops.map((op) => op.kind)).toEqual(['equal', 'delete', 'insert', 'equal']);
	});

	it('handles empty sides', () => {
		expect(diffKeys([], [])).toEqual([]);
		expect(diffKeys(['a'], []).map((op) => op.kind)).toEqual(['delete']);
		expect(diffKeys([], ['a']).map((op) => op.kind)).toEqual(['insert']);
	});

	it('is valid and minimal on random sequences', () => {
		let seed = 7;
		const random = () => {
			seed = (seed * 1103515245 + 12345) % 2147483648;
			return seed / 2147483648;
		};
		for (let round = 0; round < 300; round++) {
			const make = () =>
				Array.from({ length: Math.floor(random() * 12) }, () => 'abcd'[Math.floor(random() * 4)]!);
			const a = make();
			const b = make();
			const ops = diffKeys(a, b);
			check(a, b, ops);
			expect(ops.filter((op) => op.kind === 'equal')).toHaveLength(lcs(a, b));
		}
	});
});

describe('slideRegions', () => {
	const tokens = (...keys: string[]): Token[] =>
		keys.map((key) =>
			key === '¶'
				? { kind: 'mark', key, paragraph: { type: 'paragraph', id: 'p', runs: [] } }
				: { kind: 'text', key, pieces: [] },
		);

	it('moves an insertion to end at the paragraph mark', () => {
		const a = tokens('epsilon', '.', '¶');
		const b = tokens('epsilon', '.', 'Zeta', '.', '¶');
		const ops: DiffOp[] = [
			{ kind: 'equal', aIndex: 0, bIndex: 0 },
			{ kind: 'insert', bIndex: 1 },
			{ kind: 'insert', bIndex: 2 },
			{ kind: 'equal', aIndex: 1, bIndex: 3 },
			{ kind: 'equal', aIndex: 2, bIndex: 4 },
		];
		const slid = slideRegions(ops, a, b);
		check(
			a.map((token) => token.key),
			b.map((token) => token.key),
			slid,
		);
		expect(slid.filter((op) => op.kind === 'insert').map((op) => b[op.bIndex]!.key)).toEqual([
			'Zeta',
			'.',
		]);
	});
});

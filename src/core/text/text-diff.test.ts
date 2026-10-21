import { describe, expect, it } from 'vitest';
import { diffText, type TextChange } from './text-diff';

/** Applies the changes to `from`, which must give `to`. */
function apply(from: string, to: string, changes: readonly TextChange[]): string {
	let out = '';
	let at = 0;
	for (const change of changes) {
		out += from.slice(at, change.fromStart) + to.slice(change.toStart, change.toEnd);
		at = change.fromEnd;
	}
	return out + from.slice(at);
}

describe('diffText', () => {
	it('reports nothing for equal strings', () => {
		expect(diffText('same', 'same')).toEqual([]);
		expect(diffText('', '')).toEqual([]);
	});

	it('reports an insertion, a deletion and a replacement', () => {
		expect(diffText('ac', 'abc')).toEqual([{ fromStart: 1, fromEnd: 1, toStart: 1, toEnd: 2 }]);
		expect(diffText('abc', 'ac')).toEqual([{ fromStart: 1, fromEnd: 2, toStart: 1, toEnd: 1 }]);
		expect(diffText('a1c', 'a22c')).toEqual([{ fromStart: 1, fromEnd: 2, toStart: 1, toEnd: 3 }]);
	});

	it('keeps separate edits apart, with the unchanged text between them', () => {
		const changes = diffText('A value Z', 'B value Y');
		expect(changes).toEqual([
			{ fromStart: 0, fromEnd: 1, toStart: 0, toEnd: 1 },
			{ fromStart: 8, fromEnd: 9, toStart: 8, toEnd: 9 },
		]);
	});

	it('produces changes that rebuild the new text', () => {
		const pairs = [
			['The quick brown fox', 'A quick red fox jumps'],
			['kitten', 'sitting'],
			['abcabba', 'cbabac'],
			['', 'new'],
			['old', ''],
		] as const;
		for (const [from, to] of pairs) {
			const changes = diffText(from, to);
			expect(apply(from, to, changes)).toBe(to);
			for (const [index, change] of changes.entries()) {
				expect(change.fromEnd).toBeGreaterThanOrEqual(change.fromStart);
				expect(change.toEnd).toBeGreaterThanOrEqual(change.toStart);
				if (index) expect(change.fromStart).toBeGreaterThan(changes[index - 1]!.fromEnd);
			}
		}
	});

	it('finds a shortest edit script', () => {
		// A seeded generator, so a failure always reproduces.
		let seed = 7;
		const next = () => (seed = (seed * 48271) % 2147483647) / 2147483647;
		const word = () =>
			Array.from({ length: Math.floor(next() * 12) }, () => 'abc'[Math.floor(next() * 3)]).join('');
		const lcs = (a: string, b: string) => {
			let row = new Array<number>(b.length + 1).fill(0);
			for (const char of a) {
				const nextRow = [0];
				for (let j = 0; j < b.length; j++)
					nextRow.push(char === b[j] ? row[j]! + 1 : Math.max(row[j + 1]!, nextRow[j]!));
				row = nextRow;
			}
			return row[b.length]!;
		};
		for (let round = 0; round < 300; round++) {
			const from = word();
			const to = word();
			const changes = diffText(from, to);
			expect(apply(from, to, changes)).toBe(to);
			const edits = changes.reduce(
				(sum, change) => sum + change.fromEnd - change.fromStart + change.toEnd - change.toStart,
				0,
			);
			expect(edits, `${from} -> ${to}`).toBe(from.length + to.length - 2 * lcs(from, to));
		}
	});

	it('never starts or ends a change inside a surrogate pair', () => {
		const from = 'a\u{1F600}b';
		const to = 'a\u{1F601}b';
		const changes = diffText(from, to);
		expect(changes).toEqual([{ fromStart: 1, fromEnd: 3, toStart: 1, toEnd: 3 }]);
		expect(apply(from, to, changes)).toBe(to);
	});

	it('reports one change past the edit limit', () => {
		expect(diffText('x abcdef y', 'x uvwxyz y', 4)).toEqual([
			{ fromStart: 2, fromEnd: 8, toStart: 2, toEnd: 8 },
		]);
	});
});

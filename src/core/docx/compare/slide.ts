import type { DiffOp } from './diff';
import type { Token } from './tokens';

const SPACE = /^\s+$/u;
const SENTENCE_END = /^[.!?]$/u;

/**
 * How natural a span edge before `tokens[index]` is: paragraph marks best, then the start of a
 * sentence, then spaces, then punctuation, then the middle of a word.
 */
function boundary(tokens: Token[], index: number): number {
	const before = tokens[index - 1];
	const after = tokens[index];
	if (!before || !after) return 4;
	if (before.kind === 'mark' || after.kind === 'mark') return 5;
	if (before.kind !== 'text' || after.kind !== 'text') return 2;
	const sentence = (token: Token | undefined) =>
		token?.kind === 'text' && SENTENCE_END.test(token.key);
	if (
		SPACE.test(before.key)
			? sentence(tokens[index - 2])
			: sentence(before) && !SPACE.test(after.key)
	)
		return 4;
	if (SPACE.test(before.key) || SPACE.test(after.key)) return 3;
	const word = (token: Token) => /[\p{L}\p{N}]/u.test(token.key);
	return word(before) && word(after) ? 0 : 1;
}

/**
 * Slides each run of only insertions (or only deletions) along the equal tokens around it to the
 * position with the most natural edges, keeping the edit the same size: "epsilon.[Zeta.]" rather
 * than "epsilon[.Zeta].", as Word marks it.
 */
export function slideRegions(ops: DiffOp[], a: Token[], b: Token[]): DiffOp[] {
	const result = [...ops];
	for (let s = 0; s < result.length;) {
		if (result[s]!.kind === 'equal') {
			s++;
			continue;
		}
		let e = s;
		while (e < result.length && result[e]!.kind !== 'equal') e++;
		const kind = result[s]!.kind;
		if (result.slice(s, e).some((op) => op.kind !== kind)) {
			s = e;
			continue;
		}
		const insert = kind === 'insert';
		const tokens = insert ? b : a;
		const first = result[s]!;
		const start =
			first.kind === 'insert' ? first.bIndex : first.kind === 'delete' ? first.aIndex : 0;
		const length = e - s;
		const key = (index: number) => tokens[index]?.key;
		let left = 0;
		while (
			s - left - 1 >= 0 &&
			result[s - left - 1]!.kind === 'equal' &&
			key(start - left - 1) === key(start + length - left - 1)
		)
			left++;
		let right = 0;
		while (
			e + right < result.length &&
			result[e + right]!.kind === 'equal' &&
			key(start + length + right) === key(start + right)
		)
			right++;
		const score = (shift: number) =>
			boundary(tokens, start + shift) + boundary(tokens, start + shift + length);
		let best = 0;
		for (let shift = -left; shift <= right; shift++) if (score(shift) > score(best)) best = shift;
		if (best !== 0) {
			const window = result.slice(s - left, e + right);
			const others = window
				.filter((op): op is Extract<DiffOp, { kind: 'equal' }> => op.kind === 'equal')
				.map((op) => (insert ? op.aIndex : op.bIndex));
			const rebuilt: DiffOp[] = [];
			let other = 0;
			for (let index = start - left; index < start + length + right; index++) {
				const inRegion = index >= start + best && index < start + best + length;
				if (inRegion)
					rebuilt.push(
						insert ? { kind: 'insert', bIndex: index } : { kind: 'delete', aIndex: index },
					);
				else {
					const paired = others[other++]!;
					rebuilt.push(
						insert
							? { kind: 'equal', aIndex: paired, bIndex: index }
							: { kind: 'equal', aIndex: index, bIndex: paired },
					);
				}
			}
			result.splice(s - left, window.length, ...rebuilt);
		}
		s = e + right;
	}
	return result;
}

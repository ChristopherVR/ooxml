import { afterEach, describe, expect, it } from 'vitest';
import { chargeWork, ErrorSignal, MAX_FORMULA_WORK, resetCalcBudget } from '../values';
import { compileRegex, groupText, search } from '.';

const find = (pattern: string, text: string, ignoreCase = false) => {
	const match = search(compileRegex(pattern, ignoreCase), text, 0);
	return match ? text.slice(match[0], match[1]) : null;
};

afterEach(() => resetCalcBudget());

describe('regex matcher', () => {
	// Modifier groups are left out of the differential test: V8 applies `(?i:)` to `\w` and literals
	// outside the group. These answers are the specification's (and JavaScriptCore's).
	it('scopes modifier groups to their body', () => {
		expect(find('(?i:a)b', 'Ab')).toBe('Ab');
		expect(find('(?i:a)b', 'AB')).toBeNull();
		expect(find('(?-i:a)', 'A', true)).toBeNull();
		expect(find('A|(?i:w|)|', 'a.1')).toBe('');
		expect(find('\\w', 'ſ')).toBeNull();
		expect(find('(?i:\\w)', 'ſ')).toBe('ſ');
		expect(find('(?s:.)', '\n')).toBe('\n');
		expect(find('(?m:^b)', 'a\nb')).toBe('b');
	});

	it('clears the captures of an iteration and refuses empty iterations', () => {
		const text = 'ab';
		const match = search(compileRegex('(?:(a)|b){2}', false), text, 0);
		expect(match && groupText(text, match, 1)).toBeUndefined();
		const empty = search(compileRegex('(a?)?', false), 'b', 0);
		expect(empty && groupText('b', empty, 1)).toBeUndefined();
	});

	it('matches lookbehind right to left, as JavaScript does', () => {
		const text = '1053';
		const match = search(compileRegex('(?<=(\\d+)(\\d+))$', false), text, 0);
		expect(match && [groupText(text, match, 1), groupText(text, match, 2)]).toEqual(['1', '053']);
	});

	it('stops with #NUM! once the recalculation budget is spent', () => {
		// A backreference keeps the memo off while the matcher spends the remaining work.
		const program = compileRegex('^(a)+\\1$', false);
		resetCalcBudget();
		chargeWork(MAX_FORMULA_WORK - 1000);
		const started = performance.now();
		let thrown: unknown;
		try {
			search(program, `${'a'.repeat(256)}!`, 0);
		} catch (error) {
			thrown = error;
		}
		expect(performance.now() - started).toBeLessThan(1000);
		expect(thrown).toBeInstanceOf(ErrorSignal);
		expect((thrown as ErrorSignal).value.error).toBe('#NUM!');
	});

	it('unrolls up to a size, then counts repeats', () => {
		expect(find('(?:ab){3,5}c', 'ababababc')).toBe('ababababc');
		expect(find('(?:(?:ab){100}){100}', 'ab'.repeat(10_000))).toBe('ab'.repeat(10_000));
		expect(find('(?:a|b){20000}x', `${'ab'.repeat(10_000)}x`)).toBe(`${'ab'.repeat(10_000)}x`);
	});
});

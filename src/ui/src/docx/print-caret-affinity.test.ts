// @vitest-environment jsdom
import { expect, it } from 'vitest';
import { affinityOnLine, lineForCaret } from './print-caret-affinity';

const line = (a: number, b: number) => {
	const el = document.createElement('div');
	el.dataset.sourceStart = String(a);
	el.dataset.sourceEnd = String(b);
	return el;
};

it('draws a boundary offset at the end of the earlier line when upstream', () => {
	const block = [line(0, 6), line(6, 12)];
	expect(lineForCaret(block, 6, 'upstream')).toBe(block[0]);
	expect(lineForCaret(block, 6, 'downstream')).toBe(block[1]);
	expect(lineForCaret(block, 3, 'upstream')).toBe(block[0]);
	expect(lineForCaret(block, 9, 'upstream')).toBe(block[1]);
	expect(lineForCaret(block, 40, 'upstream')).toBeUndefined();
});

it('is upstream only at the end of a line followed by a continuation line', () => {
	const block = [line(0, 6), line(6, 12)];
	expect(affinityOnLine(block[0]!, 6, block)).toBe('upstream');
	expect(affinityOnLine(block[1]!, 6, block)).toBe('downstream');
	expect(affinityOnLine(block[1]!, 12, block)).toBe('downstream');
	expect(affinityOnLine(block[0]!, 3, block)).toBe('downstream');
});

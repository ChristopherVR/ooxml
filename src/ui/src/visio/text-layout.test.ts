import { describe, it, expect } from 'vitest';
import type { VisioTextRun } from 'ooxml-core/visio';
import { wrapText } from './text-layout.js';
const run = (text: string): VisioTextRun => ({
	text,
	fontFamily: 'Arial',
	fontSize: 1,
	color: '#000',
	bold: false,
	italic: false,
	underline: false,
});
const measure = (text: string) => Array.from(text).length;
const text = (line: { runs: VisioTextRun[] }) => line.runs.map((r) => r.text).join('');
describe('text wrapping', () => {
	it('wraps words within width and preserves rich styles', () => {
		const lines = wrapText([run('hello '), { ...run('world'), bold: true }], 6, measure);
		expect(lines.map(text)).toEqual(['hello ', 'world']);
		expect(lines[1]!.runs[0]!.bold).toBe(true);
	});
	it('honors explicit line breaks and blank lines', () => {
		expect(wrapText([run('one\r\n\r\nthree')], 20, measure).map(text)).toEqual([
			'one',
			'',
			'three',
		]);
	});
	it('splits long tokens without corrupting surrogate pairs', () => {
		expect(wrapText([run('a😀bc')], 2, measure).map(text)).toEqual(['a😀', 'bc']);
	});
	it('uses the tallest font in each line', () => {
		const lines = wrapText([run('a'), { ...run('b'), fontSize: 2 }], 10, measure);
		expect(lines[0]!.height).toBe(2.4);
	});
	it('does not mutate source runs', () => {
		const input = [run('longword')];
		wrapText(input, 3, measure);
		expect(input[0]!.text).toBe('longword');
	});
});

import { vi } from 'vitest';
import { createTextLayoutBudget, TextLayoutBudgetError } from './text-layout.js';
describe('bounded browser text work', () => {
	it('rejects a large shape before tokenization or browser measurements', () => {
		const measure = vi.fn((value: string) => value.length);
		expect(() => wrapText([run('x'.repeat(50_001))], 1, measure)).toThrow(TextLayoutBudgetError);
		expect(measure).not.toHaveBeenCalled();
	});
	it('shares a character budget across different shapes', () => {
		const budget = createTextLayoutBudget({ maxCharacters: 8, maxShapeCharacters: 8 });
		wrapText([run('abcd')], 10, measure, 0, budget);
		wrapText([run('efgh')], 10, measure, 0, budget);
		expect(budget.characters).toBe(8);
		expect(() => wrapText([run('i')], 10, measure, 0, budget)).toThrow('character budget');
	});
	it('caches repeated glyphs in long tokens instead of measuring every code point', () => {
		const measure = vi.fn((value: string) => value.length);
		const budget = createTextLayoutBudget();
		const lines = wrapText([run('a'.repeat(20_000))], 10_000, measure, 0, budget);
		expect(lines).toHaveLength(2);
		expect(lines.map(text).join('')).toHaveLength(20_000);
		expect(measure).toHaveBeenCalledTimes(4);
		expect(budget.measurements).toBe(4);
	});
	it('stops actual measure calls before exceeding the configured budget', () => {
		const measure = vi.fn((value: string) => value.length);
		const budget = createTextLayoutBudget({ maxMeasurements: 2 });
		expect(() => wrapText([run('abcd')], 2, measure, 0, budget)).toThrow('measurement budget');
		expect(measure).toHaveBeenCalledTimes(2);
		expect(budget.measurements).toBe(2);
	});
	it('enforces a smaller per-shape measurement budget without poisoning other shapes', () => {
		const budget = createTextLayoutBudget({ maxShapeMeasurements: 2 });
		expect(() => wrapText([run('abcd')], 2, measure, 0, budget)).toThrow('measurement budget');
		expect(() => wrapText([run('a')], 2, measure, 0, budget)).not.toThrow();
	});
	it('keys cached metrics by font style and the injected measurer', () => {
		const first = vi.fn((value: string) => value.length);
		const second = vi.fn((value: string) => value.length * 2);
		const budget = createTextLayoutBudget();
		wrapText([run('a')], 10, first, 0, budget);
		wrapText([{ ...run('a'), bold: true }], 10, first, 0, budget);
		const lines = wrapText([run('a')], 10, second, 0, budget);
		expect(first).toHaveBeenCalledTimes(2);
		expect(second).toHaveBeenCalledTimes(1);
		expect(lines[0]!.width).toBe(2);
	});
	it('bounds cache storage and falls back to budgeted measurement on cache overflow', () => {
		const measure = vi.fn((value: string) => value.length);
		const budget = createTextLayoutBudget({ maxCacheEntries: 1 });
		wrapText([run('a')], 10, measure, 0, budget);
		wrapText([run('a')], 10, measure, 0, budget);
		wrapText([run('b')], 10, measure, 0, budget);
		expect(measure).toHaveBeenCalledTimes(3);
	});
	it('bounds aggregate lines including empty lines', () => {
		const budget = createTextLayoutBudget({ maxLines: 2 });
		wrapText([run('\n')], 10, measure, 0, budget);
		wrapText([run('\n')], 10, measure, 0, budget);
		expect(() => wrapText([run('\n')], 10, measure, 0, budget)).toThrow('line budget');
		expect(budget.lines).toBe(2);
	});
	it('rejects hostile font specifications and invalid metrics', () => {
		expect(() => wrapText([{ ...run('a'), fontFamily: 'x'.repeat(1025) }], 1, measure)).toThrow(
			'font specification',
		);
		expect(() => wrapText([run('a')], 1, () => NaN)).toThrow('invalid width');
		expect(() => wrapText([run('a')], 1, () => -1)).toThrow('invalid width');
	});
	it('rejects invalid budget configuration and retains zero-width no-wrap behavior', () => {
		expect(() => createTextLayoutBudget({ maxMeasurements: 0 })).toThrow('positive safe integers');
		expect(wrapText([run('abcdef')], 0, measure).map(text)).toEqual(['abcdef']);
	});
});
it('coalesces identical word/space styles instead of generating one SVG fragment per token', () => {
	const budget = createTextLayoutBudget();
	const content = 'a '.repeat(2000);
	const lines = wrapText([run(content)], 100_000, measure, 0, budget);
	expect(lines).toHaveLength(1);
	expect(lines[0]!.runs).toHaveLength(1);
	expect(lines[0]!.runs[0]!.text).toBe(content);
	expect(budget.fragments).toBe(1);
});
it('bounds distinct styled fragments across shapes without dropping style differences', () => {
	const budget = createTextLayoutBudget({ maxFragments: 2 });
	const lines = wrapText([run('a'), { ...run('b'), color: '#f00' }], 100, measure, 0, budget);
	expect(lines[0]!.runs.map((value) => value.color)).toEqual(['#000', '#f00']);
	expect(() => wrapText([run('c')], 100, measure, 0, budget)).toThrow('styled-fragment budget');
	expect(budget.fragments).toBe(2);
});

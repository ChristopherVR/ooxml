import { describe, it, expect } from 'vitest';
import { demoDocument } from './demo-document.js';
import { layoutParagraphs } from './paragraph-layout.js';
import type { VisioParagraph } from 'ooxml-core/visio';
const paragraph: VisioParagraph = {
	start: 0,
	end: 11,
	horizontalAlign: 'left',
	indentLeft: 1,
	indentRight: 0,
	indentFirst: 1,
	spaceBefore: 0.5,
	spaceAfter: 0.3,
	lineSpacing: { kind: 'exact', value: 0.4 },
	direction: 'ltr',
};
describe('paragraph layout', () => {
	it('applies first-line indentation, paragraph spacing and exact line height', () => {
		const text = structuredClone(demoDocument.pages[0]!.shapes[0]!.text);
		text.plainText = 'hello world';
		text.width = 8;
		text.fontSize = 0.2;
		text.paragraphs = [paragraph];
		const result = layoutParagraphs(text, (s) => s.length);
		expect(result.lines).toHaveLength(2);
		expect(result.lines[0]!.x).toBe(2.06);
		expect(result.lines[1]!.x).toBe(1.06);
		expect(result.lines[0]!.height).toBe(0.4);
		expect(result.height).toBeCloseTo(1.6);
	});
	it('preserves mixed run styles, direction and bullets in layout', () => {
		const text = structuredClone(demoDocument.pages[0]!.shapes[0]!.text);
		text.plainText = 'abc';
		text.width = 8;
		text.paragraphs = [
			{
				...paragraph,
				end: 3,
				direction: 'rtl',
				horizontalAlign: 'right',
				bullet: { text: '•', fontFamily: 'Arial', fontSize: 0.2, offset: 0.3 },
			},
		];
		const result = layoutParagraphs(text, (s) => s.length);
		expect(result.lines[0]!.direction).toBe('rtl');
		expect(result.lines[0]!.bullet?.text).toBe('•');
		expect(result.lines[0]!.align).toBe('right');
	});
});

import { vi } from 'vitest';
import { createTextLayoutBudget } from './text-layout.js';
describe('indexed and bounded paragraph layout', () => {
	it('indexes runs once rather than scanning every preceding run for each paragraph', () => {
		const text = structuredClone(demoDocument.pages[0]!.shapes[0]!.text);
		const count = 2500;
		let reads = 0;
		text.plainText = 'x\n'.repeat(count);
		text.width = 100;
		text.runs = Array.from({ length: count }, () => ({
			get text() {
				++reads;
				return 'x\n';
			},
			fontFamily: 'Arial',
			fontSize: 0.2,
			color: '#000',
			bold: false,
			italic: false,
			underline: false,
		}));
		text.paragraphs = Array.from({ length: count }, (_, i) => ({
			...paragraph,
			start: i * 2,
			end: i * 2 + 1,
			indentLeft: 0,
			indentFirst: 0,
		}));
		const budget = createTextLayoutBudget();
		const result = layoutParagraphs(text, (value) => value.length, budget);
		expect(result.lines).toHaveLength(count);
		expect(reads).toBeLessThan(count * 4);
		expect(budget.characters).toBe(count * 2);
		expect(budget.measurements).toBe(1);
	});
	it('rejects overlapping, reversed or nonintegral ranges before measuring', () => {
		const text = structuredClone(demoDocument.pages[0]!.shapes[0]!.text);
		text.plainText = 'abcdef';
		const measure = vi.fn((value: string) => value.length);
		for (const ranges of [
			[
				{ start: 0, end: 4 },
				{ start: 3, end: 6 },
			],
			[
				{ start: 4, end: 6 },
				{ start: 0, end: 2 },
			],
			[{ start: 0.5, end: 3 }],
		]) {
			text.paragraphs = ranges.map((range) => ({ ...paragraph, ...range }));
			expect(() => layoutParagraphs(text, measure)).toThrow('ordered, non-overlapping');
		}
		expect(measure).not.toHaveBeenCalled();
	});
	it('counts paragraph separators and blank lines against shared budgets', () => {
		const text = structuredClone(demoDocument.pages[0]!.shapes[0]!.text);
		text.plainText = '\n\n\n';
		text.runs = [];
		text.paragraphs = [0, 1, 2].map((start) => ({ ...paragraph, start, end: start }));
		const budget = createTextLayoutBudget({ maxLines: 2 });
		expect(() => layoutParagraphs(text, (value) => value.length, budget)).toThrow('line budget');
		expect(budget.characters).toBe(3);
		expect(budget.lines).toBe(2);
	});
	it('applies one per-shape measurement limit across all its paragraphs', () => {
		const text = structuredClone(demoDocument.pages[0]!.shapes[0]!.text);
		text.plainText = 'a\nb\nc';
		text.runs = [];
		text.paragraphs = [0, 2, 4].map((start) => ({ ...paragraph, start, end: start + 1 }));
		const measure = vi.fn((value: string) => value.length);
		expect(() =>
			layoutParagraphs(text, measure, createTextLayoutBudget({ maxShapeMeasurements: 2 })),
		).toThrow('measurement budget');
		expect(measure).toHaveBeenCalledTimes(2);
	});
	it('preserves offsets across run boundaries and empty runs', () => {
		const text = structuredClone(demoDocument.pages[0]!.shapes[0]!.text);
		text.plainText = 'abcd';
		text.width = 100;
		text.runs = ['', 'ab', '', 'cd', ''].map((value) => ({
			text: value,
			fontFamily: 'Arial',
			fontSize: 0.2,
			color: '#000',
			bold: false,
			italic: false,
			underline: false,
		}));
		text.paragraphs = [{ ...paragraph, start: 1, end: 3 }];
		const result = layoutParagraphs(text, (value) => value.length);
		expect(result.lines[0]!.runs.map((run) => run.text).join('')).toBe('bc');
	});
});

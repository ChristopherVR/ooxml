import { describe, expect, it } from 'vitest';
import { compactPageText, statusLayout } from './status-layout';

describe('statusLayout', () => {
	it('shows everything before the bar is measured', () => {
		expect(statusLayout(0)).toEqual({
			words: true,
			language: true,
			views: true,
			compactPage: false,
			stats: Infinity,
		});
	});
	it('shows everything on a desktop bar', () => {
		const layout = statusLayout(1366);
		expect(layout).toMatchObject({ words: true, language: true, views: true });
		expect(layout.stats).toBe(Infinity);
	});
	it('drops the language, then aggregates, on a tablet', () => {
		expect(statusLayout(600)).toEqual({
			words: true,
			language: false,
			views: true,
			compactPage: false,
			stats: 2,
		});
	});
	it('keeps only the essentials on a phone', () => {
		expect(statusLayout(375)).toEqual({
			words: false,
			language: false,
			views: false,
			compactPage: true,
			stats: 0,
		});
	});
	it('keeps one aggregate and the word count on a wide phone in landscape', () => {
		expect(statusLayout(500)).toEqual({
			words: true,
			language: false,
			views: true,
			compactPage: true,
			stats: 1,
		});
	});
});

describe('compactPageText', () => {
	it('keeps the two page numbers', () => {
		expect(compactPageText('Page 2 of 5')).toBe('2/5');
		expect(compactPageText('第 2 页，共 5 页')).toBe('2/5');
		expect(compactPageText('')).toBe('');
	});
});

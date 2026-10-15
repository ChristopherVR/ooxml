import { describe, expect, it } from 'vitest';
import { statusLayout } from './status-layout';

describe('statusLayout', () => {
	it('shows everything before the bar is measured', () => {
		expect(statusLayout(0)).toEqual({
			words: true,
			language: true,
			views: true,
			stats: Infinity,
		});
	});
	it('shows everything on a desktop bar', () => {
		const layout = statusLayout(1366);
		expect(layout).toMatchObject({ words: true, language: true, views: true });
		expect(layout.stats).toBe(Infinity);
	});
	it('drops the language, then aggregates, on a tablet', () => {
		expect(statusLayout(600)).toEqual({ words: true, language: false, views: true, stats: 2 });
	});
	it('keeps only the essentials on a phone', () => {
		expect(statusLayout(375)).toEqual({ words: false, language: false, views: false, stats: 0 });
	});
	it('keeps one aggregate and the word count on a wide phone in landscape', () => {
		expect(statusLayout(500)).toEqual({ words: true, language: false, views: true, stats: 1 });
	});
});

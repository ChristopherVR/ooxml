import { describe, expect, it } from 'vitest';

import { linkSpec } from './part-specs';

describe('linkSpec.decode', () => {
	it('keeps the link fields a document holds', () => {
		expect(
			linkSpec.decode({
				r: 'A1',
				target: 'https://example.com/',
				location: 'Sheet2!A1',
				tooltip: 'Open',
				display: 'Example',
			}),
		).toEqual({
			range: { start: { row: 0, col: 0 }, end: { row: 0, col: 0 } },
			target: 'https://example.com/',
			location: 'Sheet2!A1',
			tooltip: 'Open',
			display: 'Example',
		});
	});

	it('drops a script target and any field that is not text', () => {
		const link = linkSpec.decode({
			r: 'B2',
			target: 'javascript:alert(1)',
			tooltip: { toString: 'x' },
			extra: '<b>',
		});
		expect(link).toEqual({ range: { start: { row: 1, col: 1 }, end: { row: 1, col: 1 } } });
	});
});

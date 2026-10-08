import { describe, expect, it } from 'vitest';
import { sourcePaint } from './source-paint';
import type { Cells } from './sheet';

describe('cached source paint metadata', () => {
	it.each([25, 40, 254])(
		'retains unsupported pattern %i for an honest current-value label',
		(pattern) => {
			const cells: Cells = new Map([['FillPattern', { value: String(pattern) }]]);
			expect(sourcePaint(cells, () => '#ffffff').fillPatternIndex).toBe(pattern);
		},
	);
	it('omits unresolved and error caches instead of substituting rendered values', () => {
		const cells: Cells = new Map([
			['FillPattern', { value: 'Themed' }],
			['FillForegndTrans', { value: 'Themed' }],
			['FillBkgndTrans', { value: '0.5', error: '1' }],
			['LineColorTrans', { value: '2' }],
			['FillBkgnd', { value: 'Themed' }],
		]);
		expect(sourcePaint(cells, () => '')).toEqual({});
	});
});

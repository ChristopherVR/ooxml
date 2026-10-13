import { expect, it } from 'vitest';
import { nestedFrameBorderWidth } from './nested-frame-profile';
import type { LayoutTable, LayoutCellBorders } from './input';
const table = (borders?: LayoutCellBorders): LayoutTable => ({
	kind: 'table',
	id: 'nested',
	rows: [{ cells: [{ paragraphs: [], ...(borders ? { borders } : {}) }] }],
});
const edge = { widthPx: 4 / 3, style: 'solid' as const, color: '#FF0000' };
it('admits only measured uniform single-line border occupancy profiles', () => {
	const borders = { top: edge, right: edge, bottom: edge, left: edge };
	expect(nestedFrameBorderWidth(table())).toBe(0);
	expect(nestedFrameBorderWidth(table(borders))).toBeCloseTo(4 / 3);
	expect(
		nestedFrameBorderWidth(table({ ...borders, right: { ...edge, widthPx: 2 } })),
	).toBeUndefined();
	expect(
		nestedFrameBorderWidth(table({ ...borders, top: { ...edge, style: 'dashed' } })),
	).toBeUndefined();
	expect(nestedFrameBorderWidth(table({ top: edge }))).toBeUndefined();
});

import { expect, it } from 'vitest';
import { nestedFrameProfile } from './nested-frame-profile';
import type { LayoutTable, LayoutCellBorders } from './input';
const table = (borders?: LayoutCellBorders): LayoutTable => ({
	kind: 'table',
	id: 'nested',
	rows: [{ cells: [{ paragraphs: [], ...(borders ? { borders } : {}) }] }],
});
const edge = { widthPx: 4 / 3, style: 'solid' as const, color: '#FF0000' };
it('admits only measured shared-edge-consistent single-line border occupancy profiles', () => {
	const borders = { top: edge, right: edge, bottom: edge, left: edge };
	expect(nestedFrameProfile(table())).toEqual({ leftPx: 0, rightPx: 0, bottomPx: 0 });
	expect(nestedFrameProfile(table(borders))).toEqual({
		leftPx: 4 / 3,
		rightPx: 4 / 3,
		bottomPx: 4 / 3,
	});
	expect(nestedFrameProfile(table({ ...borders, right: { ...edge, widthPx: 2 } }))).toEqual({
		leftPx: 4 / 3,
		rightPx: 2,
		bottomPx: 4 / 3,
	});
	expect(
		nestedFrameProfile(table({ ...borders, top: { ...edge, style: 'dashed' } })),
	).toBeUndefined();
	expect(nestedFrameProfile(table({ top: edge }))).toBeUndefined();
});

it('refuses conflicting shared edges and changing row grids', () => {
	const borders = { top: edge, right: edge, bottom: edge, left: edge };
	const cell = { paragraphs: [], xPx: 0, widthPx: 80, borders };
	const good: LayoutTable = {
		kind: 'table',
		id: 'nested',
		rows: [{ cells: [cell, { ...cell, xPx: 80 }] }],
	};
	expect(nestedFrameProfile(good)).toBeDefined();
	expect(
		nestedFrameProfile({
			...good,
			rows: [
				{
					cells: [
						cell,
						{ ...cell, xPx: 80, borders: { ...borders, left: { ...edge, widthPx: 2 } } },
					],
				},
			],
		}),
	).toBeUndefined();
	expect(
		nestedFrameProfile({ ...good, rows: [good.rows[0]!, { cells: [{ ...cell, widthPx: 160 }] }] }),
	).toBeUndefined();
	expect(
		nestedFrameProfile({ ...good, rows: [good.rows[0]!, { cells: [cell, { ...cell, xPx: 81 }] }] }),
	).toBeUndefined();
});

it('refuses border centers that would produce a negative painted cell width', () => {
	const borders = { top: edge, bottom: edge, left: edge, right: { ...edge, widthPx: 200 } };
	const following = { ...borders, left: borders.right, right: edge };
	expect(
		nestedFrameProfile({
			kind: 'table',
			id: 'nested',
			widthPx: 160,
			rows: [
				{
					cells: [
						{ paragraphs: [], widthPx: 80, borders },
						{ paragraphs: [], widthPx: 80, xPx: 80, borders: following },
					],
				},
			],
		}),
	).toBeUndefined();
});

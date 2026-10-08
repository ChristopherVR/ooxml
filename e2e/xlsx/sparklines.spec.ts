import { readFileSync } from 'node:fs';
import { expect, test, type Page } from '@playwright/test';
import {
	FRAMEWORKS,
	editor,
	goToCell,
	grid,
	openLanding,
	pageErrors,
	ribbon,
	typeInActiveCell,
} from './helpers';

// Authored by Excel 16 (src/core/xlsx/__fixtures__/generate-excel-sparklines.ps1). Sheet1 holds
// A1:E4 and one sparkline group per host cell:
//   F1 line   1, 3, -2, 5, 4    markers, high, low and negative points
//   F2 column 2, 4, _, 1, 3     first and last points, the empty C2 leaves a gap
//   F3 win/loss 1, -1, 1, 1, -1 negative points
//   F4 line   2, 8, 6, 3, 7     manual scale 0..10 with the horizontal axis shown
const fixture = readFileSync(
	new URL('../../src/core/xlsx/__fixtures__/excel-sparklines.xlsx', import.meta.url),
);
const SERIES = '#376092';
const POINT = '#D00000';

interface Paint {
	width: number;
	height: number;
	cellWidth: number;
	polylines: string[];
	strokes: string[];
	circles: string[];
	rects: { y: number; height: number; fill: string }[];
	axis: { y: number; stroke: string }[];
}

/** The sparkline SVGs of the grid, top to bottom, as painted. */
function sparks(page: Page): Promise<Paint[]> {
	return editor(page).evaluate((node) => {
		const root = node.shadowRoot!;
		const hosts = [...root.querySelectorAll<HTMLElement>('.xg-c')]
			.filter((cell) => cell.querySelector(':scope > .xg-spark'))
			.sort((a, b) => Number.parseFloat(a.style.top) - Number.parseFloat(b.style.top));
		return hosts.map((cell) => {
			const svg = cell.querySelector(':scope > .xg-spark')!;
			const all = (tag: string) => [...svg.querySelectorAll(tag)];
			return {
				width: Number(svg.getAttribute('width')),
				height: Number(svg.getAttribute('height')),
				cellWidth: Number.parseFloat(cell.style.width),
				polylines: all('polyline').map((p) => p.getAttribute('points') ?? ''),
				strokes: all('polyline').map((p) => p.getAttribute('stroke') ?? ''),
				circles: all('circle').map((c) => c.getAttribute('fill') ?? ''),
				rects: all('rect').map((r) => ({
					y: Number(r.getAttribute('y')),
					height: Number(r.getAttribute('height')),
					fill: r.getAttribute('fill') ?? '',
				})),
				axis: all('line').map((l) => ({
					y: Number(l.getAttribute('y1')),
					stroke: l.getAttribute('stroke') ?? '',
				})),
			};
		});
	});
}

async function openSparklines(page: Page, framework: string) {
	await openLanding(page, framework);
	await page.locator('#landing-file').setInputFiles({
		name: 'excel-sparklines.xlsx',
		mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
		buffer: fixture,
	});
	await expect(grid(page)).toBeVisible();
	await expect.poll(async () => (await sparks(page)).length).toBe(4);
}

for (const framework of FRAMEWORKS)
	test(`${framework}: Excel sparklines paint in their host cells`, async ({ page }) => {
		const errors = pageErrors(page);
		await openSparklines(page, framework);
		const [line, column, winLoss, axisLine] = await sparks(page);

		// F1: one unbroken line through five points, every point marked in the point colour.
		expect(line!.polylines).toHaveLength(1);
		expect(line!.polylines[0]!.split(' ')).toHaveLength(5);
		expect(line!.strokes).toEqual([SERIES]);
		expect(line!.circles).toEqual(Array(5).fill(POINT));
		expect(line!.rects).toEqual([]);
		expect(line!.axis).toEqual([]);

		// F2: four columns (C2 is empty), the first and last in the point colour.
		expect(column!.rects.map((r) => r.fill)).toEqual([POINT, SERIES, SERIES, POINT]);
		expect(column!.polylines).toEqual([]);

		// F3: five win/loss bars of one height, wins above the middle and losses below it.
		expect(winLoss!.rects.map((r) => r.fill)).toEqual([SERIES, POINT, SERIES, SERIES, POINT]);
		const middle = winLoss!.height / 2;
		const heights = new Set(winLoss!.rects.map((r) => r.height));
		expect(heights.size).toBe(1);
		for (const [index, rect] of winLoss!.rects.entries())
			if (index === 1 || index === 4) expect(rect.y).toBeGreaterThanOrEqual(middle - 0.5);
			else expect(rect.y + rect.height).toBeLessThanOrEqual(middle + 0.5);

		// F4: the manual 0..10 scale puts the axis at zero, the bottom of the cell.
		expect(axisLine!.axis).toHaveLength(1);
		expect(axisLine!.axis[0]!.stroke).toBe('#000000');
		expect(axisLine!.axis[0]!.y).toBeGreaterThan(axisLine!.height / 2);
		expect(axisLine!.polylines).toHaveLength(1);
		expect(axisLine!.circles).toEqual([]);

		for (const paint of [line, column, winLoss, axisLine])
			expect(paint!.width).toBe(paint!.cellWidth);
		expect(errors).toEqual([]);
	});

test('editing the data range repaints the sparkline and a wider column widens it', async ({
	page,
}) => {
	const errors = pageErrors(page);
	await openSparklines(page, 'vanilla');
	const before = (await sparks(page))[0]!;

	// D1 holds the line's high point; a new high moves every point of the scaled line.
	await goToCell(page, 'D1');
	await typeInActiveCell(page, '50');
	await expect
		.poll(async () => (await sparks(page))[0]!.polylines[0])
		.not.toBe(before.polylines[0]);

	await goToCell(page, 'F1');
	await ribbon(page).getByRole('button', { name: 'Format', exact: true }).click();
	await editor(page).getByRole('menuitem', { name: 'Column Width...', exact: true }).click();
	const dialog = editor(page).locator('[data-dialog="column-width"]');
	await dialog.getByLabel('Column width:', { exact: true }).fill('30');
	await dialog.getByRole('button', { name: 'OK', exact: true }).click();
	await expect.poll(async () => (await sparks(page))[0]!.width).toBeGreaterThan(before.width);
	const after = await sparks(page);
	for (const paint of after) expect(paint.width).toBe(paint.cellWidth);
	expect(errors).toEqual([]);
});

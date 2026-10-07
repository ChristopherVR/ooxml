import { expect, test } from '@playwright/test';
import { readFileSync } from 'node:fs';
import { createRequire } from 'node:module';
import { editor, goToCell, grid, newWorkbook } from './helpers';

// These tests run from the viewer workspace, which owns the JSZip dependency.
const JSZip: typeof import('jszip') = createRequire(
	new URL('../../viewers/xlsx/package.json', import.meta.url),
)('jszip');
const native = JSON.parse(
	readFileSync(
		new URL('../../src/core/xlsx/edit/__fixtures__/excel-databar-clipboard.json', import.meta.url),
		'utf8',
	),
) as {
	cases: { variant: number; before: string }[];
};

test('native solid RTL data bars retain colors and borders through clipboard paste and undo', async ({
	page,
}) => {
	await page.context().grantPermissions(['clipboard-read', 'clipboard-write']);
	await newWorkbook(page);
	const zip = await JSZip.loadAsync(
		readFileSync(new URL('./support/excel-smartart.xlsx', import.meta.url)),
	);
	zip.file('xl/worksheets/sheet1.xml', native.cases.find((record) => record.variant === 1)!.before);
	const bytes = await zip.generateAsync({ type: 'uint8array' });
	await editor(page).evaluate(
		async (node, data) => {
			await (node as unknown as { load(bytes: Uint8Array, name: string): Promise<void> }).load(
				new Uint8Array(data),
				'native-data-bars.xlsx',
			);
		},
		[...bytes],
	);
	const bars = grid(page).locator('.xg-c:not([hidden]) .xg-db');
	await expect(bars).toHaveCount(5);
	// All five original bars use the same explicit direction and hide their values.
	for (let i = 0; i < 5; i++) {
		await expect(bars.nth(i)).toHaveCSS('background-image', 'none');
		await expect(bars.nth(i).locator('..').locator('.xg-db-axis')).toHaveCSS(
			'border-left-color',
			'rgb(255, 0, 255)',
		);
	}
	await expect(bars.nth(1)).toHaveCSS('background-color', 'rgb(0, 0, 255)');
	await expect(bars.nth(1)).toHaveCSS('border-top-color', 'rgb(255, 255, 0)');
	await expect(bars.nth(4)).toHaveCSS('background-color', 'rgb(0, 255, 0)');
	await expect(bars.nth(4)).toHaveCSS('border-top-color', 'rgb(255, 0, 0)');
	await goToCell(page, 'A2');
	await page.keyboard.press('Control+C');
	await goToCell(page, 'C2');
	await page.keyboard.press('Control+V');
	await expect(bars).toHaveCount(6);
	await expect
		.poll(() =>
			bars.evaluateAll(
				(nodes) =>
					nodes.filter((node) => getComputedStyle(node).backgroundColor === 'rgb(0, 0, 255)')
						.length,
			),
		)
		.toBe(3);
	await grid(page).focus();
	await page.keyboard.press('Control+Z');
	await expect(bars).toHaveCount(5);
	await page.keyboard.press('Control+Y');
	await expect(bars).toHaveCount(6);
});

import { expect, test } from '@playwright/test';
import { editor, editorProperty, openSample, pageErrors } from './helpers';

test('shared chart gallery supports keyboard selection, grouping edits and save/reopen', async ({
	page,
}) => {
	const errors = pageErrors(page);
	await openSample(page);
	await editor(page).getByRole('combobox', { name: 'Name Box', exact: true }).fill('A3:C8');
	await page.keyboard.press('Enter');
	await editor(page).getByRole('tab', { name: 'Insert', exact: true }).click();
	await editor(page)
		.getByRole('button', { name: 'Recommended Charts', exact: true })
		.first()
		.click();
	const dialog = editor(page).locator('[data-dialog="insert-chart"]');
	const gallery = dialog.locator('office-ui-gallery[mode="panel"]');
	await expect(gallery.locator('.tile svg')).toHaveCount(8);
	await page.setViewportSize({ width: 390, height: 844 });
	const bounds = await gallery.boundingBox();
	expect(bounds!.x).toBeGreaterThanOrEqual(0);
	expect(bounds!.x + bounds!.width).toBeLessThanOrEqual(390);
	await page.setViewportSize({ width: 1280, height: 900 });
	await expect(gallery.getByRole('button', { name: 'Column', exact: true })).toBeFocused();
	await gallery.getByRole('button', { name: 'Column', exact: true }).press('ArrowDown');
	await expect(gallery.getByRole('button', { name: 'Area', exact: true })).toBeFocused();
	await expect(gallery.getByRole('button', { name: 'Area', exact: true })).toHaveAttribute(
		'aria-pressed',
		'true',
	);
	await gallery.getByRole('button', { name: 'Bar', exact: true }).click();
	await dialog.getByLabel('Grouping', { exact: true }).selectOption('stacked');
	await dialog.getByLabel('Chart title', { exact: true }).fill('Grouped sales');
	await dialog.getByRole('button', { name: 'OK', exact: true }).click();
	await editor(page).getByRole('tab', { name: 'Chart Design', exact: true }).click();
	await editor(page)
		.getByRole('button', { name: 'Change Chart Type', exact: true })
		.first()
		.click();
	await expect(dialog.getByLabel('Grouping', { exact: true })).toHaveValue('stacked');
	await dialog.getByLabel('Grouping', { exact: true }).selectOption('percentStacked');
	await dialog.getByRole('button', { name: 'Column', exact: true }).click();
	await dialog.getByRole('button', { name: 'OK', exact: true }).click();
	type Book = {
		sheets: {
			drawings: { kind: string; chartType?: string; grouping?: string; title?: string }[];
		}[];
	};
	const chart = async () =>
		(await editorProperty<Book>(page, 'workbook')).sheets[0]!.drawings.find(
			(drawing) => drawing.title === 'Grouped sales',
		);
	expect(await chart()).toMatchObject({ chartType: 'column', grouping: 'percentStacked' });
	await editor(page).evaluate((node) => (node as unknown as { undo(): void }).undo());
	expect(await chart()).toMatchObject({ chartType: 'bar', grouping: 'stacked' });
	await editor(page).evaluate((node) => (node as unknown as { redo(): void }).redo());
	const bytes = await editor(page).evaluate(async (node) => [
		...new Uint8Array(
			await (await (node as unknown as { save(): Promise<Blob> }).save()).arrayBuffer(),
		),
	]);
	await editor(page).evaluate(async (node, data) => {
		await (node as unknown as { load(bytes: Uint8Array, name: string): Promise<void> }).load(
			new Uint8Array(data),
			'chart-roundtrip.xlsx',
		);
	}, bytes);
	expect(await chart()).toMatchObject({ chartType: 'column', grouping: 'percentStacked' });
	expect(errors).toEqual([]);
});

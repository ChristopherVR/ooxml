import { expect, test } from '@playwright/test';
import { FRAMEWORKS, editor, editorProperty, openSample, pageErrors } from './helpers';

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

for (const framework of FRAMEWORKS)
	test(`Change Colors uses the shared gallery in ${framework}`, async ({ page }) => {
		const errors = pageErrors(page);
		await openSample(page, framework);
		await editor(page).getByRole('combobox', { name: 'Name Box', exact: true }).fill('A3:C8');
		await page.keyboard.press('Enter');
		await editor(page).getByRole('tab', { name: 'Insert', exact: true }).click();
		await editor(page)
			.getByRole('button', { name: 'Recommended Charts', exact: true })
			.first()
			.click();
		const dialog = editor(page).locator('[data-dialog="insert-chart"]');
		await dialog.getByLabel('Chart title', { exact: true }).fill('Palette sales');
		await dialog.getByRole('button', { name: 'OK', exact: true }).click();
		await editor(page).getByRole('tab', { name: 'Chart Design', exact: true }).click();
		const trigger = editor(page).locator('[data-gallery="chart.colors"]');
		await trigger.press('ArrowDown');
		const popup = editor(page).locator('[data-gallery-popup="chart.colors"]');
		await expect(popup.locator('.tile svg')).toHaveCount(17);
		await expect(
			popup.getByRole('button', { name: 'Colorful Palette 1', exact: true }),
		).toBeFocused();
		await page.keyboard.press('End');
		await expect(
			popup.getByRole('button', { name: 'Monochromatic Palette 13', exact: true }),
		).toBeFocused();
		await page.keyboard.press('Home');
		await page.keyboard.press('ArrowDown');
		await expect(
			popup.getByRole('button', { name: 'Colorful Palette 2', exact: true }),
		).toBeFocused();
		// At phone width the ribbon may fold Chart Styles into a dropdown (it depends on the
		// platform's font metrics), hiding the open gallery with it. Let the ribbon refit, then
		// reopen the gallery, from the folded group when there is one, before measuring it.
		await page.keyboard.press('Escape');
		await page.setViewportSize({ width: 390, height: 844 });
		await page.evaluate(
			() => new Promise((done) => requestAnimationFrame(() => requestAnimationFrame(done))),
		);
		const folded = editor(page).getByRole('button', { name: 'Chart Styles', exact: true });
		if (await folded.isVisible()) await folded.click();
		await trigger.click();
		await expect(popup).toBeVisible();
		const bounds = await popup.boundingBox();
		expect(bounds!.x).toBeGreaterThanOrEqual(0);
		expect(bounds!.x + bounds!.width).toBeLessThanOrEqual(390);
		await page.keyboard.press('Escape');
		await page.setViewportSize({ width: 1280, height: 900 });
		await expect(trigger).toBeVisible();
		await trigger.click();
		await popup.getByRole('button', { name: 'Monochromatic Palette 2', exact: true }).click();
		type Book = { sheets: { drawings: { title?: string; colorPalette?: number }[] }[] };
		const chart = async () =>
			(await editorProperty<Book>(page, 'workbook')).sheets[0]!.drawings.find(
				(d) => d.title === 'Palette sales',
			);
		expect((await chart())?.colorPalette).toBe(15);
		await trigger.click();
		await expect(
			popup.getByRole('button', { name: 'Monochromatic Palette 2', exact: true }),
		).toHaveAttribute('aria-pressed', 'true');
		await page.keyboard.press('Escape');
		await expect(trigger).toBeFocused();
		await editor(page).evaluate((node) => (node as unknown as { undo(): void }).undo());
		expect((await chart())?.colorPalette).toBeUndefined();
		await editor(page).evaluate((node) => (node as unknown as { redo(): void }).redo());
		const bytes = await editor(page).evaluate(async (node) => [
			...new Uint8Array(
				await (await (node as unknown as { save(): Promise<Blob> }).save()).arrayBuffer(),
			),
		]);
		await editor(page).evaluate(
			async (node, data) =>
				(node as unknown as { load(bytes: Uint8Array, name: string): Promise<void> }).load(
					new Uint8Array(data),
					'palette-roundtrip.xlsx',
				),
			bytes,
		);
		expect((await chart())?.colorPalette).toBe(15);
		expect(errors).toEqual([]);
	});

for (const framework of FRAMEWORKS)
	test(`Chart Styles and Quick Layout use the shared galleries in ${framework}`, async ({
		page,
	}) => {
		const errors = pageErrors(page);
		await openSample(page, framework);
		await editor(page).getByRole('combobox', { name: 'Name Box', exact: true }).fill('A3:C8');
		await page.keyboard.press('Enter');
		await editor(page).getByRole('tab', { name: 'Insert', exact: true }).click();
		await editor(page)
			.getByRole('button', { name: 'Recommended Charts', exact: true })
			.first()
			.click();
		const dialog = editor(page).locator('[data-dialog="insert-chart"]');
		await dialog.getByLabel('Chart title', { exact: true }).fill('Design sales');
		await dialog.getByRole('button', { name: 'OK', exact: true }).click();
		await editor(page).getByRole('tab', { name: 'Chart Design', exact: true }).click();
		type Series = { drawingColor?: { value: string }; dataLabels?: { chart?: object } };
		type Book = {
			sheets: {
				drawings: {
					kind: string;
					title?: string;
					quickLayout?: number;
					showLegend?: boolean;
					legendPosition?: string;
					series: Series[];
				}[];
			}[];
		};
		const drawings = async () => (await editorProperty<Book>(page, 'workbook')).sheets[0]!.drawings;
		// Layout 4 has no title, so the chart is found by its position once inserted.
		const index = (await drawings()).findIndex((d) => d.title === 'Design sales');
		expect(index).toBeGreaterThanOrEqual(0);
		const chart = async () => (await drawings())[index];
		const styles = editor(page).locator('[data-gallery="chart.styles"]');
		await styles.click();
		const stylePopup = editor(page).locator('[data-gallery-popup="chart.styles"]');
		await expect(stylePopup.locator('.tile svg')).toHaveCount(6);
		await stylePopup.getByRole('button', { name: 'Monochrome', exact: true }).click();
		const pinned = (await chart())!.series.map((series) => series.drawingColor?.value);
		expect(pinned.every(Boolean)).toBe(true);
		await styles.click();
		await expect(
			stylePopup.getByRole('button', { name: 'Monochrome', exact: true }),
		).toHaveAttribute('aria-pressed', 'true');
		await page.keyboard.press('Escape');
		const layouts = editor(page).locator('[data-gallery="chart.quick-layout"]');
		await layouts.click();
		const layoutPopup = editor(page).locator('[data-gallery-popup="chart.quick-layout"]');
		await expect(layoutPopup.locator('.tile svg')).toHaveCount(11);
		await layoutPopup.getByRole('button', { name: 'Layout 4', exact: true }).click();
		expect((await chart())?.title).toBeUndefined();
		expect(await chart()).toMatchObject({
			quickLayout: 4,
			showLegend: true,
			legendPosition: 'b',
		});
		expect((await chart())!.series[0]?.dataLabels?.chart).toBeTruthy();
		const bytes = await editor(page).evaluate(async (node) => [
			...new Uint8Array(
				await (await (node as unknown as { save(): Promise<Blob> }).save()).arrayBuffer(),
			),
		]);
		await editor(page).evaluate(
			async (node, data) =>
				(node as unknown as { load(bytes: Uint8Array, name: string): Promise<void> }).load(
					new Uint8Array(data),
					'design-roundtrip.xlsx',
				),
			bytes,
		);
		const back = (await chart())!;
		expect(back.series[0]?.dataLabels?.chart).toBeTruthy();
		expect(back.series.map((series) => series.drawingColor?.value)).toEqual(pinned);
		expect(errors).toEqual([]);
	});

import { readFileSync } from 'node:fs';
import path from 'node:path';
import { expect, test } from '@playwright/test';
import JSZip from 'jszip';
import { dialogByHeading, reopenDocument } from './helpers';
const source = readFileSync(
	path.join(import.meta.dirname, '../../src/core/docx/__fixtures__/chart-families.docx'),
);
for (const framework of ['vanilla', 'react', 'vue', 'angular', 'svelte', 'solid'])
	test(`${framework}: ChartEx title and family edits save and reload`, async ({ page }) => {
		test.setTimeout(120000);
		const errors: string[] = [];
		page.on('pageerror', (e) => errors.push(e.message));
		await page.setViewportSize({ width: 1500, height: 1000 });
		await page.goto(`/?framework=${framework}`);
		await reopenDocument(page, 'chartex.docx', source);
		const editor = page.locator('docx-editor');
		const chart = editor
			.locator('.ProseMirror [data-docx-chart="1"]')
			.filter({ has: page.locator('svg text').filter({ hasText: 'Cash flow' }) });
		await chart.click();
		await editor.getByRole('tab', { name: 'Chart Design', exact: true }).click();
		await editor.getByRole('button', { name: 'Edit Data', exact: true }).click();
		const dialog = dialogByHeading(editor, 'Edit Data');
		const options = dialog.locator('office-ui-chart-ex-options');
		await options.getByLabel('Chart title', { exact: true }).fill('Browser ChartEx');
		await options.getByLabel('Chart title', { exact: true }).press('Tab');
		await options.getByLabel('Connector lines', { exact: true }).uncheck();
		await dialog.getByRole('button', { name: 'OK', exact: true }).click();
		await editor.getByRole('button', { name: 'Change Chart Type', exact: true }).click();
		const typeDialog = dialogByHeading(editor, 'Change Chart Type');
		await expect(typeDialog.locator('office-ui-chart-ex-options')).toBeVisible();
		await typeDialog.getByRole('button', { name: 'OK', exact: true }).click();
		await editor.getByRole('button', { name: 'Format Selection', exact: true }).click();
		const pane = editor.locator('office-ui-chart-format-pane');
		await pane.locator('[data-chart-format-target]').selectOption('series:0');
		await pane.getByRole('radio', { name: 'Pattern fill', exact: true }).click();
		await pane
			.locator('[data-chart-format-field="fill.patternPreset"] select')
			.selectOption('cross');
		await pane.locator('[data-chart-format-close]').click();
		const bytes = Uint8Array.from(
			await editor.evaluate(async (element) =>
				Array.from(await (element as unknown as { saveBytes(): Promise<Uint8Array> }).saveBytes()),
			),
		);
		const zip = await JSZip.loadAsync(bytes);
		const xml = await zip.file('word/charts/chartEx1.xml')!.async('string');
		expect(xml).toContain('Browser ChartEx');
		expect(xml).toContain('connectorLines="0"');
		expect(xml).toContain('prst="cross"');
		await reopenDocument(page, 'chartex-saved.docx', bytes);
		await expect(
			editor
				.locator('.ProseMirror [data-docx-chart="1"] svg text')
				.filter({ hasText: 'Browser ChartEx' }),
		).toBeVisible();
		expect(errors).toEqual([]);
	});

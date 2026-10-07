import { fileURLToPath } from 'node:url';
import { test, expect } from '@playwright/test';
import { fileInput } from './helpers';
import type { DocxEditorElement } from '../../viewers/docx/packages/web-component/src';

const fixture = (area: string, name: string, suffix: string) =>
	fileURLToPath(
		new URL(`../../src/core/docx/__fixtures__/${area}/${name}-${suffix}.docx`, import.meta.url),
	);
for (const framework of ['vanilla', 'react', 'vue', 'angular', 'svelte', 'solid'])
	for (const [area, name] of [
		['review-formatting', 'bold'],
		['review-formatting', 'multiple'],
		['review-paragraph-formatting', 'multiple'],
	] as const)
		test(`${framework}: Original Print Layout matches native ${area} ${name} before formatting`, async ({
			page,
		}) => {
			await page.setViewportSize({ width: 2400, height: 1000 });
			const errors: string[] = [];
			page.on('pageerror', (error) => errors.push(error.message));
			await page.goto(`/?framework=${framework}`);
			const input = await fileInput(page);
			await input.setInputFiles(fixture(area, name, 'before'));
			const editor = page.locator('docx-editor');
			await editor.getByRole('tab', { name: 'View', exact: true }).click();
			const layout = editor.getByRole('combobox', { name: 'Layout view', exact: true });
			await layout.selectOption('print');
			const columns = editor.locator('.dve-canvas > .dve-print-pages .dve-print-column');
			await expect(columns.first()).toBeVisible();
			await page.evaluate(() => document.fonts.ready);
			const appearance = () =>
				columns.evaluateAll((elements) => elements.map((element) => element.innerHTML));
			const before = await appearance();
			await input.setInputFiles(fixture(area, name, 'tracked'));
			await expect.poll(appearance).not.toEqual(before);
			const current = await appearance();
			const source = await editor.evaluate(
				(element) => (element as DocxEditorElement).documentModel,
			);
			await editor.getByRole('tab', { name: 'Review', exact: true }).click();
			const mode = editor.getByRole('combobox', { name: 'Display for review', exact: true });
			await mode.selectOption('original', { force: true });
			await expect.poll(appearance).toEqual(before);
			expect(
				await editor.evaluate((element) => (element as DocxEditorElement).documentModel),
			).toEqual(source);
			await mode.selectOption('final', { force: true });
			await expect.poll(appearance).toEqual(current);
			await mode.selectOption('original', { force: true });
			await expect.poll(appearance).toEqual(before);
			await columns.first().locator('.dve-print-line').first().click();
			await expect(editor.locator('.dve-paper')).toBeVisible();
			expect(
				await editor.evaluate((element) => (element as DocxEditorElement).documentModel),
			).toEqual(source);
			expect(errors).toEqual([]);
		});

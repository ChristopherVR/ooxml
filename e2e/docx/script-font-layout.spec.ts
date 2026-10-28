import { expect, test } from '@playwright/test';
import type { DocxEditorElement } from '../../viewers/docx/packages/web-component/src';
import { newDocument } from './helpers';

for (const framework of ['vanilla', 'react', 'vue', 'angular', 'svelte', 'solid']) {
	test(`${framework}: print layout selects script fonts and retains source offsets`, async ({
		page,
	}) => {
		const errors: string[] = [];
		page.on('pageerror', (error) => errors.push(error.message));
		await page.goto(`/?framework=${framework}`);
		await newDocument(page);
		const editor = page.locator('docx-editor');
		await editor.evaluate((element) => {
			const host = element as DocxEditorElement;
			host.documentModel = {
				...host.documentModel!,
				characterStyles: {
					docDefaults: {
						fontFamilyEastAsia: 'SimSun',
						fontFamilyComplexScript: 'Traditional Arabic',
					},
					styles: {},
					warnings: [],
				},
				blocks: [
					{ type: 'paragraph', id: 'scripts', runs: [{ text: 'A中عB', fontFamily: 'Calibri' }] },
				],
			};
		});
		await editor.getByRole('tab', { name: 'View', exact: true }).click();
		await editor.getByRole('button', { name: 'Print Layout', exact: true }).click();
		const line = editor.locator(
			'.dve-canvas > .dve-print-pages .dve-print-line[data-block-id="scripts"]',
		);
		await expect(line).toHaveText('A中عB');
		const fragments = await line.locator('span[data-source-start]').evaluateAll((spans) =>
			spans.map((span) => ({
				text: span.textContent,
				family: (span as HTMLElement).style.fontFamily,
				start: span.getAttribute('data-source-start'),
				end: span.getAttribute('data-source-end'),
			})),
		);
		expect(fragments.map(({ text, start, end }) => [text, start, end])).toEqual([
			['A', '0', '1'],
			['中', '1', '2'],
			['ع', '2', '3'],
			['B', '3', '4'],
		]);
		expect(fragments[1]?.family).toContain('SimSun');
		expect(fragments[2]?.family).toContain('Traditional Arabic');
		expect(errors).toEqual([]);
	});
}

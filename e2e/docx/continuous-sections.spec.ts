import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { expect, test } from '@playwright/test';
import { fileInput } from './helpers';

const fixture = (name: string) =>
	new URL(`../../src/core/docx/layout/fixtures/continuous-sections/${name}`, import.meta.url);
const evidence = JSON.parse(await readFile(fixture('evidence.json'), 'utf8')) as {
	cases: {
		name: string;
		pages: number;
		positions: { text: string; page: number; xPt: number; yPt: number }[];
	}[];
};

for (const framework of ['vanilla', 'react', 'vue', 'angular', 'svelte', 'solid'])
	for (const reference of evidence.cases)
		test(`${framework}: native continuous ${reference.name} layout`, async ({ page }) => {
			const errors: string[] = [];
			page.on('pageerror', (error) => errors.push(error.message));
			page.on('dialog', (dialog) => void dialog.accept());
			await page.goto(`/?framework=${framework}`);
			await (await fileInput(page)).setInputFiles(fileURLToPath(fixture(`${reference.name}.docx`)));
			const editor = page.locator('docx-editor');
			await expect(editor.locator('.ProseMirror')).toContainText('After2');
			await editor.getByRole('tab', { name: 'View', exact: true }).click();
			await editor.getByRole('button', { name: 'Print Layout', exact: true }).click();
			// Page thumbnails also contain print-page clones. Inspect the main print surface.
			const sheets = editor.locator('.dve-canvas > .dve-print-pages .dve-print-page');
			await expect(sheets).toHaveCount(reference.pages);
			const positions = await sheets.evaluateAll((elements) =>
				elements.flatMap((sheet, index) =>
					Array.from(sheet.querySelectorAll<HTMLElement>('.dve-print-line')).map((line) => {
						const column = line.closest<HTMLElement>('.dve-print-column')!;
						const block = line.closest<HTMLElement>('.dve-print-block')!;
						return {
							text: line.textContent,
							page: index + 1,
							xPt: parseFloat(column.style.left) * 0.75,
							yPt:
								(parseFloat(column.style.top) +
									parseFloat(block.style.top) +
									parseFloat(line.style.top)) *
								0.75,
						};
					}),
				),
			);
			for (const position of reference.positions.filter((entry) => entry.text))
				expect(positions.find((entry) => entry.text === position.text)).toEqual({
					text: position.text,
					page: position.page,
					xPt: position.xPt,
					yPt: position.yPt,
				});
			expect(errors).toEqual([]);
		});

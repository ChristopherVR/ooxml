import { fileURLToPath } from 'node:url';
import { expect, test } from '@playwright/test';
import { fileInput } from './helpers';

// Word's own file of text boxes (`src/core/docx/layout/fixtures/text-boxes/README.md` lists what
// Word 16 draws): Print Layout lays their paragraphs out with the page engine.
const fixture = fileURLToPath(
	new URL('../../src/core/docx/layout/fixtures/text-boxes/word-saved.docx', import.meta.url),
);

for (const framework of ['vanilla', 'react'])
	test(`${framework}: text boxes are laid out by the page engine in Print Layout`, async ({
		page,
	}) => {
		const errors: string[] = [];
		page.on('pageerror', (error) => errors.push(error.message));
		page.on('dialog', (dialog) => void dialog.accept());
		await page.goto(`/?framework=${framework}`);
		await (await fileInput(page)).setInputFiles(fixture);
		const editor = page.locator('docx-editor');
		await expect(editor.locator('.ProseMirror')).toContainText('Text box layout');
		await editor.getByRole('tab', { name: 'View', exact: true }).click();
		await editor.getByRole('button', { name: 'Print Layout', exact: true }).click();
		const sheet = editor.locator('.dve-canvas > .dve-print-pages .dve-print-page').first();
		const boxes = sheet.locator('.dve-print-float');
		await expect(boxes).toHaveCount(10);
		await expect(sheet.locator('.dve-text-box-content')).toHaveCount(0);
		const lines = (index: number) =>
			boxes.nth(index).locator('.dve-print-text-box .dve-print-line');
		// Lines break inside the text rectangle, as Word breaks them.
		await expect(lines(0).first()).toHaveText('Wrapped word1 word2');
		await expect(lines(0).nth(1)).toHaveText('word3 word4 word5');
		// The overflow is clipped at the text rectangle.
		const clipped = await boxes
			.nth(0)
			.locator('.dve-print-text-box')
			.evaluate((layer) => getComputedStyle(layer).overflow);
		expect(clipped).toBe('hidden');
		// The linked box continues the first box's text.
		const firstBox = (await lines(6).allTextContents()).join(' ');
		const secondBox = await lines(7).first().textContent();
		const last = Number(/word(\d+)\s*$/.exec(firstBox)?.[1]);
		expect(secondBox).toMatch(new RegExp(`^word${last + 1} `));
		// Vertical text is turned: the first line of a vert box is at its right.
		const vertical = await boxes.nth(2).evaluate((box) => {
			const frame = box.getBoundingClientRect();
			const first = box.querySelector('.dve-print-text-box .dve-print-line')!;
			const line = first.getBoundingClientRect();
			return { right: frame.right - line.right, tall: line.height > line.width };
		});
		expect(vertical.tall).toBe(true);
		expect(vertical.right).toBeLessThan(20);
		// spAutoFit grows the shape to its text (stored 0.4 inch, 38 px).
		const autofit = await boxes.nth(4).evaluate((box) => box.getBoundingClientRect().height);
		expect(autofit).toBeGreaterThan(80);
		expect(errors).toEqual([]);
	});

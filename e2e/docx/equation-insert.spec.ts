import { expect, test } from '@playwright/test';
import JSZip from 'jszip';
import { dialogByHeading, newDocument } from './helpers';

const savedXml = async (page: import('@playwright/test').Page) => {
	const bytes = await page
		.locator('docx-editor')
		.evaluate(async (element) =>
			Array.from(await (element as unknown as { saveBytes(): Promise<Uint8Array> }).saveBytes()),
		);
	return (await JSZip.loadAsync(new Uint8Array(bytes))).file('word/document.xml')!.async('string');
};

test('Insert > Equation inserts a gallery equation, edits it from LaTeX and draws it on the page', async ({
	page,
}) => {
	await page.setViewportSize({ width: 1700, height: 900 });
	await page.goto('/?framework=vanilla');
	await newDocument(page);
	const editor = page.locator('docx-editor');
	await editor.locator('.ProseMirror').click();
	await editor.getByRole('tab', { name: 'Insert', exact: true }).click();
	await editor
		.getByRole('combobox', { name: 'Equation', exact: true })
		.selectOption({ label: 'Quadratic Formula' });
	const equation = editor.locator('.ProseMirror [data-docx-equation]');
	await expect(equation).toHaveCount(1);
	await expect(equation.locator('msqrt')).toHaveCount(1);
	await expect(equation.locator('math')).toHaveAttribute('display', 'block');

	// Double-clicking opens the LaTeX box with live preview.
	await equation.dblclick();
	const dialog = dialogByHeading(editor, 'Equation');
	await expect(dialog).toBeVisible();
	const source = dialog.getByRole('textbox', { name: 'LaTeX equation', exact: true });
	await expect(source).toHaveValue(/frac/);
	await source.fill('\\frac{a}{b} + \\sum_{i=1}^{n} i');
	await expect(dialog.locator('.dve-equation-preview mfrac')).toHaveCount(1);
	await source.fill('\\frac{a');
	await expect(dialog.getByRole('button', { name: 'OK', exact: true })).toBeDisabled();
	await source.fill('\\frac{a}{b} + \\sum_{i=1}^{n} i');
	await dialog.getByRole('button', { name: 'OK', exact: true }).click();
	await expect(equation.locator('munderover, msubsup')).toHaveCount(1);

	// Print Layout draws it as a sized object instead of a "[Equation]" marker.
	await editor.getByRole('tab', { name: 'View', exact: true }).click();
	await editor.getByRole('button', { name: 'Print Layout', exact: true }).click();
	const printed = editor.locator('.dve-print-pages [data-docx-chart="1"] svg');
	await expect(printed).toHaveCount(1);
	expect((await editor.locator('.dve-print-pages').allTextContents()).join('')).not.toContain(
		'[Equation]',
	);
	const box = await printed.boundingBox();
	expect(box!.width).toBeGreaterThan(40);
	expect(box!.height).toBeGreaterThan(20);

	const xml = await savedXml(page);
	expect(xml).toContain('<m:oMathPara');
	expect(xml).toContain('<m:nary>');
	expect(xml).toContain('<m:f>');
});

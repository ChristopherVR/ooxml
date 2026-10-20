import { expect, test } from '@playwright/test';
import { openSample, FRAMEWORKS } from './helpers';
type XlsxEditorElement = HTMLElementTagNameMap['xlsx-editor'];

for (const framework of FRAMEWORKS) {
	test(`${framework}: host controls edit and save a standalone spreadsheet grid`, async ({
		page,
	}) => {
		await openSample(page, framework);
		await page.evaluate(() => {
			const old = document.querySelector('xlsx-editor') as XlsxEditorElement;
			const surface = document.createElement('xlsx-editor') as XlsxEditorElement;
			surface.id = 'composed';
			surface.surfaceOnly = true;
			surface.workbook = old.workbook;
			surface.style.height = '600px';
			const bold = document.createElement('button');
			bold.textContent = 'Custom bold';
			bold.onclick = () => {
				void surface.getToolbarProps().onBold();
			};
			old.replaceWith(bold, surface);
			surface.select('B20');
			surface.focusGrid();
		});
		const surface = page.locator('#composed');
		await expect(
			surface.locator(
				'office-ui-ribbon, office-ui-title-bar, office-ui-backstage, [part="formula-bar"], [part="sheet-tabs"]',
			),
		).toHaveCount(0);
		await page.keyboard.type('123');
		await expect
			.poll(() =>
				page.evaluate(
					() =>
						(document.querySelector('#composed') as XlsxEditorElement).getToolbarProps().formula,
				),
			)
			.toBe('123');
		await page.getByRole('button', { name: 'Custom bold' }).click();
		await expect
			.poll(() =>
				page.evaluate(
					() => (document.querySelector('#composed') as XlsxEditorElement).getToolbarProps().bold,
				),
			)
			.toBe(true);
		const result = await page.evaluate(async () => {
			const editor = document.querySelector('#composed') as XlsxEditorElement;
			const before = editor.getSelection();
			const typed = editor.getToolbarProps().formula;
			const drafted = editor.getToolbarProps().onFormulaInput('=1+2');
			const draft = editor.getToolbarProps().formula;
			const committed = editor.getToolbarProps().commitEdit();
			const bytes = await editor.saveBytes();
			await editor.load(bytes);
			editor.select('B20');
			editor.readOnly = true;
			return {
				before,
				typed,
				drafted,
				draft,
				committed,
				loaded: !editor.loading,
				permitted: await editor.getToolbarProps().onItalic(),
				formula: editor.getToolbarProps().formula,
			};
		});
		expect(result.before).toBe('B20');
		expect(result.loaded).toBe(true);
		expect(result.permitted).toBe(false);
		expect(result.typed).toBe('123');
		expect(result.drafted).toBe(true);
		expect(result.draft).toBe('=1+2');
		expect(result.committed).toBe(true);
		expect(result.formula).toBe('=1+2');
	});
}

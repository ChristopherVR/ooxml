import { expect, test, type Locator } from '@playwright/test';
import JSZip from 'jszip';
import { savedDocument } from './saved-document';
import { newDocument, reveal } from './helpers';

const FRAMEWORKS = ['vanilla', 'react', 'vue', 'angular', 'svelte', 'solid'] as const;

const tab = (editor: Locator, name: string) => editor.getByRole('tab', { name, exact: true });

for (const framework of FRAMEWORKS) {
	test(`${framework}: Alt+= opens the Equation tab, typing builds up and structures insert`, async ({
		page,
	}) => {
		await page.setViewportSize({ width: 1700, height: 900 });
		await page.goto(`/?framework=${framework}`);
		await newDocument(page);
		const editor = page.locator('docx-editor');
		await editor.locator('.ProseMirror').click();
		await expect(tab(editor, 'Equation')).toHaveCount(0);

		await page.keyboard.press('Alt+=');
		await expect(tab(editor, 'Equation')).toHaveAttribute('aria-selected', 'true');
		const input = editor.getByRole('textbox', { name: 'Equation editor', exact: true });
		await expect(input).toBeFocused();
		await expect(input).toHaveAttribute('placeholder', 'Type equation here.');

		// Linear input builds up, and math AutoCorrect turns \alpha into its letter on space.
		await input.pressSequentially('x=\\alpha ');
		await expect(input).toHaveValue('x=α');
		await input.pressSequentially('+');
		const panel = editor.locator('[data-ribbon-tab="equation"]');
		const fraction = panel.getByRole('button', { name: 'Fraction', exact: true });
		await reveal(editor, fraction);
		await fraction.click();
		const gallery = editor.locator('.dve-equation-gallery');
		await expect(gallery.getByRole('menuitem')).not.toHaveCount(0);
		await gallery.getByRole('menuitem', { name: 'Stacked Fraction', exact: true }).click();
		await expect(input).toHaveValue('x=α+⬚/⬚');
		await input.pressSequentially('1');
		await input.press('Tab');
		await input.pressSequentially('2');
		await expect(input).toHaveValue('x=α+1/2');
		const equation = editor.locator('.ProseMirror [data-docx-equation]');
		await expect(equation.locator('mfrac')).toHaveCount(1);

		// Symbols insert at the caret, from their category.
		const symbols = panel.getByRole('button', { name: 'Symbols', exact: true });
		await reveal(editor, symbols);
		await symbols.click();
		await editor
			.locator('.dve-equation-gallery')
			.getByRole('combobox', { name: 'Symbol category' })
			.selectOption({ label: 'Arrows' });
		await editor.locator('.dve-equation-gallery').getByRole('menuitem', { name: '⇒' }).click();
		await expect(input).toHaveValue('x=α+1/2⇒');

		const { xml, bytes } = await savedDocument(page, `ui-equation-tab-${framework}.docx`);
		expect(xml).toContain('<m:oMathPara');
		expect(xml).toMatch(/<m:f><m:num><m:r><m:t>1<\/m:t>/);
		expect(xml).toContain('<m:t>α</m:t>');
		expect((await JSZip.loadAsync(bytes)).file('word/document.xml')).not.toBeNull();

		// Leaving the equation hides the tab.
		await input.press('ArrowRight');
		await expect(tab(editor, 'Equation')).toHaveCount(0);
	});
}

test('editing an existing equation, Linear and Professional, LaTeX input, Ink Equation disabled', async ({
	page,
}) => {
	await page.setViewportSize({ width: 1700, height: 900 });
	await page.goto('/?framework=vanilla');
	await newDocument(page);
	const editor = page.locator('docx-editor');
	await editor.locator('.ProseMirror').click();
	await editor.getByRole('tab', { name: 'Insert', exact: true }).click();
	const picker = editor.getByRole('combobox', { name: 'Equation', exact: true });
	await reveal(editor, picker);
	await picker.selectOption({ label: 'Pythagorean Theorem' });
	const equation = editor.locator('.ProseMirror [data-docx-equation]');
	await expect(equation.locator('msup')).toHaveCount(3);

	// Clicking the equation puts the caret in it, in its linear text.
	await equation.click();
	const input = editor.getByRole('textbox', { name: 'Equation editor', exact: true });
	await expect(input).toHaveValue('a^2+b^2=c^2');
	await input.press('End');
	await input.pressSequentially('+d_i');
	await expect(equation.locator('msub')).toHaveCount(1);

	const panel = editor.locator('[data-ribbon-tab="equation"]');
	const ink = panel.getByRole('button', { name: 'Ink Equation', exact: true });
	await expect(ink).toBeDisabled();
	await expect(ink).toHaveAttribute('title', /handwriting recognition/);

	const convert = panel.getByRole('combobox', { name: 'Convert', exact: true });
	await reveal(editor, convert);
	await convert.selectOption({ label: 'Current - Linear' });
	await expect(equation.locator('msup')).toHaveCount(0);
	let { xml } = await savedDocument(page, 'ui-equation-linear.docx');
	expect(xml).toContain('<m:t>a^2+b^2=c^2+d_i</m:t>');
	await convert.selectOption({ label: 'Current - Professional' });
	await expect(equation.locator('msup')).toHaveCount(3);

	const latex = panel.getByRole('button', { name: 'LaTeX', exact: true });
	await reveal(editor, latex);
	await latex.click();
	await expect(latex).toHaveAttribute('aria-pressed', 'true');
	await expect(input).toHaveValue(/a\^\{2\}/);
	await input.fill('\\sqrt{x}');
	await expect(equation.locator('msqrt')).toHaveCount(1);
	({ xml } = await savedDocument(page, 'ui-equation-latex.docx'));
	expect(xml).toMatch(/<m:rad><m:radPr><m:degHide m:val="1"\/><\/m:radPr><m:deg\/><m:e>/);

	// Print Layout draws the equation and selects it on a click.
	await input.press('End');
	await input.press('ArrowRight');
	await expect(tab(editor, 'Equation')).toHaveCount(0);
	await editor.getByRole('tab', { name: 'View', exact: true }).click();
	await editor.getByRole('button', { name: 'Print Layout', exact: true }).click();
	const printed = editor.locator('.dve-print-pages [data-docx-chart="1"]');
	await expect(printed).toHaveCount(1);
	await printed.click();
	await expect(tab(editor, 'Equation')).toHaveAttribute('aria-selected', 'true');
	await expect(input).toHaveValue('\\sqrt{x}');
});

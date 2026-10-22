import { expect, test } from '@playwright/test';
import { clickInEditor, openSample, waitForEditorFocus } from './helpers';
type DocxEditorElement = HTMLElementTagNameMap['docx-editor'];

for (const framework of ['vanilla', 'react', 'vue', 'angular', 'svelte', 'solid']) {
	test(`${framework}: host controls edit a standalone Word surface`, async ({ page }) => {
		await openSample(page, framework);
		await page.evaluate(() => {
			const old = document.querySelector('docx-editor') as DocxEditorElement;
			const surface = document.createElement('docx-editor') as DocxEditorElement;
			surface.id = 'composed';
			surface.surfaceOnly = true;
			surface.documentModel = old.documentModel;
			surface.style.height = '600px';
			const bold = document.createElement('button');
			bold.textContent = 'Custom bold';
			bold.onclick = () => surface.getToolbarProps().onBold();
			old.replaceWith(bold, surface);
		});
		const surface = page.locator('#composed');
		await expect(
			surface.locator(
				'office-ui-ribbon, office-ui-title-bar, office-ui-backstage, [role="dialog"]',
			),
		).toHaveCount(0);
		const text = surface.locator('.ProseMirror');
		await clickInEditor(text.locator('p').last());
		await page.keyboard.press('End');
		// The editor reads the caret End moved only while it has focus, which the button takes next.
		await waitForEditorFocus(text);
		await page.getByRole('button', { name: 'Custom bold' }).click();
		await page.keyboard.type(' Composed text.');
		await expect(text.locator('strong').filter({ hasText: 'Composed text.' })).toHaveCount(1);
		const result = await page.evaluate(async () => {
			const editor = document.querySelector('#composed') as DocxEditorElement;
			const bytes = await editor.saveBytes();
			const dirty = editor.dirty;
			await editor.load(bytes);
			editor.readOnly = true;
			return { dirty, loaded: !editor.loading, permitted: editor.getToolbarProps().onItalic() };
		});
		expect(result).toEqual({ dirty: true, loaded: true, permitted: false });
		await expect(text.locator('strong').filter({ hasText: 'Composed text.' })).toHaveCount(1);
	});
}

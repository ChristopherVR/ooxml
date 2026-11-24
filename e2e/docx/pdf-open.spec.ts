import { expect, test } from '@playwright/test';
import { readFile } from 'node:fs/promises';
import { fileInput, fileNameLabel } from './helpers';

// A PDF that Word 16 exported (src/core/docx/pdf-import/__fixtures__, see generate.mjs there).
const rapport = () =>
	readFile(new URL('../../src/core/docx/pdf-import/__fixtures__/rapport.pdf', import.meta.url));

test('File > Open converts a PDF into an editable document, as Word does (#319)', async ({
	page,
}) => {
	page.on('dialog', (dialog) => void dialog.accept());
	await page.goto('/?framework=vanilla');
	await (
		await fileInput(page)
	).setInputFiles({ name: 'rapport.pdf', mimeType: 'application/pdf', buffer: await rapport() });
	const editor = page.locator('docx-editor');
	await expect(fileNameLabel(page)).toHaveText('rapport.docx');
	const surface = editor.locator('.ProseMirror');
	await expect(surface).toContainText('Rapport annuel');
	await expect(surface).toContainText('Un passage en gras, un en italique');
	// The running head went to the header, not the body.
	await expect(surface).not.toContainText('Société Exemple – Rapport annuel 2026');
	// The text is editable like any document's.
	await surface.getByText('Fin du rapport.').click();
	await page.keyboard.press('End');
	await page.keyboard.type(' Ajouté.');
	await expect(surface).toContainText('Fin du rapport. Ajouté.');
});

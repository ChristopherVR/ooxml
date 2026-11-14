import { expect, test, type Locator, type Page } from '@playwright/test';
import JSZip from 'jszip';
import { newDocument } from './helpers';

/** A 40x20 PNG drawn by the browser, as the bytes a dropped file would carry. */
async function pngBase64(page: Page): Promise<string> {
	return page.evaluate(() => {
		const canvas = document.createElement('canvas');
		canvas.width = 40;
		canvas.height = 20;
		canvas.getContext('2d')!.fillRect(0, 0, 40, 20);
		return canvas.toDataURL('image/png').split(',')[1]!;
	});
}

async function savedXml(editor: Locator): Promise<string> {
	const bytes = Buffer.from(
		await editor.evaluate(async (node) =>
			Array.from(await (node as unknown as { saveBytes(): Promise<Uint8Array> }).saveBytes()),
		),
	);
	return (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string');
}

/** Drops `name` (a PNG) on the canvas at a viewport point, as a file dragged from the desktop. */
async function dropPicture(editor: Locator, data: string, x: number, y: number, name = 'dot.png') {
	await editor.evaluate(
		(node, { data, x, y, name }) => {
			const bytes = Uint8Array.from(atob(data), (char) => char.charCodeAt(0));
			const transfer = new DataTransfer();
			transfer.items.add(new File([bytes], name, { type: 'image/png' }));
			const canvas = node.shadowRoot!.querySelector('.dve-canvas')!;
			const init = {
				dataTransfer: transfer,
				clientX: x,
				clientY: y,
				bubbles: true,
				composed: true,
			};
			const target = node.shadowRoot!.elementFromPoint(x, y) ?? canvas;
			target.dispatchEvent(new DragEvent('dragover', { ...init, cancelable: true }));
			target.dispatchEvent(new DragEvent('drop', { ...init, cancelable: true }));
		},
		{ data, x, y, name },
	);
}

async function documentWithPicture(page: Page) {
	await page.goto('/');
	await newDocument(page);
	const editor = page.locator('docx-editor');
	const surface = editor.locator('.ProseMirror');
	await surface.click();
	await page.keyboard.type('Before after');
	const data = await pngBase64(page);
	const line = editor.getByText('Before after').last();
	const box = (await line.boundingBox())!;
	// Between "Before" and " after": about half of the text's width.
	await dropPicture(editor, data, box.x + box.width * 0.5, box.y + box.height / 2);
	await expect(surface.locator('img[data-docx-image]')).toHaveCount(1);
	return { editor, surface };
}

test('a picture file dropped on the page is inserted there as an inline picture', async ({
	page,
}) => {
	const errors: string[] = [];
	page.on('pageerror', (error) => errors.push(error.message));
	const { editor, surface } = await documentWithPicture(page);
	await expect(surface.locator('img[data-docx-image]')).toHaveAttribute('alt', 'dot');
	const text = await surface
		.locator('p')
		.first()
		.evaluate((p) => {
			const parts: string[] = [];
			const walk = (node: Node) => {
				if (node instanceof HTMLImageElement) parts.push('[picture]');
				else if (node.nodeType === Node.TEXT_NODE) parts.push(node.textContent ?? '');
				else node.childNodes.forEach(walk);
			};
			walk(p);
			return parts.join('');
		});
	// Dropped between the words, not appended at the end.
	expect(text).toMatch(/^Bef.*\[picture\].*after$/);
	expect(text.replace('[picture]', '')).toBe('Before after');
	const xml = await savedXml(editor);
	expect(xml).toContain('<wp:inline');
	expect(xml).toContain('descr="dot"');
	expect(errors).toEqual([]);
});

test('in Print Layout a dropped picture lands at the drop point and can be selected', async ({
	page,
}) => {
	await page.goto('/');
	await newDocument(page);
	const editor = page.locator('docx-editor');
	await editor.locator('.ProseMirror').click();
	await page.keyboard.type('Before after');
	await editor.getByRole('tab', { name: 'View', exact: true }).click();
	await editor.getByRole('button', { name: 'Print Layout', exact: true }).click();
	// The line spans the column; the text fragment is just the words.
	const line = editor.locator('.dve-print-line', { hasText: 'Before after' }).first();
	const box = (await line.locator('[data-source-start]').first().boundingBox())!;
	await dropPicture(editor, await pngBase64(page), box.x + box.width * 0.5, box.y + box.height / 2);
	await expect(editor.locator('.dve-print-page .dve-print-picture')).toHaveCount(1);
	const xml = await savedXml(editor);
	const body = xml.replace(/<w:rPr>.*?<\/w:rPr>/g, '');
	expect(body).not.toMatch(/after<\/w:t><\/w:r><w:r><w:drawing>/);
	expect(body).toMatch(/<w:t[^>]*>Bef[^<]*<\/w:t><\/w:r><w:r><w:drawing>/);
	await editor.locator('.dve-print-page .dve-print-picture').click();
	await expect(editor.getByRole('tab', { name: 'Picture Format', exact: true })).toBeVisible();
});

test('Picture Format wraps and crops the picture, and the menu offers Wrap Text', async ({
	page,
}) => {
	const { editor, surface } = await documentWithPicture(page);
	await surface.locator('img[data-docx-image]').click();
	const tab = editor.getByRole('tab', { name: 'Picture Format', exact: true });
	await expect(tab).toBeVisible();
	await tab.click();
	const panel = editor.locator('[data-ribbon-tab="picture-format"]');
	await panel.getByRole('combobox', { name: 'Wrap text' }).selectOption('square');
	expect(await savedXml(editor)).toContain('<wp:wrapSquare wrapText="bothSides"/>');

	await surface.locator('.dve-picture img').first().click({ button: 'right' });
	const menu = editor.locator('office-ui-context-menu');
	await expect(menu.getByRole('menuitemcheckbox', { name: 'Square' })).toHaveAttribute(
		'aria-checked',
		'true',
	);
	await menu.getByRole('menuitemcheckbox', { name: 'Top and Bottom' }).click();
	expect(await savedXml(editor)).toContain('<wp:wrapTopAndBottom/>');

	await surface.locator('.dve-picture img').first().click();
	await panel.getByRole('button', { name: 'Crop', exact: true }).click();
	const dialog = editor.locator('.dve-picture-dialog');
	await dialog.locator('.dve-picture-crop').getByLabel('Left').fill('25');
	await expect(dialog.getByLabel('Width')).toHaveValue('30');
	await dialog.getByRole('button', { name: 'OK' }).click();
	const xml = await savedXml(editor);
	expect(xml).toContain('<a:srcRect l="25000"/>');
	expect(xml).toMatch(/<wp:extent cx="285750" cy="190500"\/>/);
});

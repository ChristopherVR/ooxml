import { expect, test } from '@playwright/test';
import { readFile } from 'node:fs/promises';
import { dialogByHeading, fileInput } from './helpers';
import { savedDocument } from './saved-document';

test('Word arranges a floating shape from its Layout ribbon and saves the edit', async ({
	page,
}) => {
	await page.setViewportSize({ width: 1800, height: 950 });
	await page.goto('/?framework=vanilla');
	await (
		await fileInput(page)
	).setInputFiles({
		name: 'shapes.docx',
		mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
		buffer: await readFile(new URL('./support/shapes.docx', import.meta.url)),
	});
	const editor = page.locator('docx-editor');
	const shape = editor.locator('.ProseMirror .dve-shape').first();
	await expect(shape).toBeVisible();
	await editor.getByRole('tab', { name: 'Layout', exact: true }).click();
	const layout = editor.getByRole('tabpanel', { name: 'Layout', exact: true });
	const altText = editor.getByRole('button', { name: 'Alt Text', exact: true });
	const unfoldArrange = async () => {
		const folded = layout.getByRole('button', { name: 'Arrange', exact: true });
		if ((await folded.isVisible()) && !(await altText.isVisible())) await folded.click();
	};
	await unfoldArrange();
	await editor.getByRole('button', { name: 'Selection Pane', exact: true }).click();
	const pane = editor.getByRole('complementary', { name: 'Selection Pane' });
	await expect(pane.getByRole('listitem')).toHaveCount(4);
	await pane.getByRole('listitem').first().getByRole('button').click();
	await expect(pane.getByRole('button', { pressed: true })).toHaveCount(1);
	await pane.getByRole('button', { name: 'Close', exact: true }).click();
	await editor.getByRole('tab', { name: 'Layout', exact: true }).click();
	await unfoldArrange();
	const position = editor.getByRole('combobox', { name: 'Position', exact: true });
	await position.press('Enter');
	await editor.getByRole('menuitemradio', { name: /Bottom Right/ }).click();
	await unfoldArrange();
	await editor.getByRole('combobox', { name: 'Rotate', exact: true }).selectOption('right');
	await unfoldArrange();
	await altText.click();
	const alt = dialogByHeading(editor, 'Alt Text');
	await alt.getByRole('textbox', { name: 'Description', exact: true }).fill('Arrangement sample');
	await alt.getByRole('button', { name: 'OK', exact: true }).click();
	const { xml } = await savedDocument(page, 'word-arrange.docx');
	expect(xml).toContain('<wp:positionH relativeFrom="margin"><wp:align>right</wp:align>');
	expect(xml).toContain('<wp:positionV relativeFrom="margin"><wp:align>bottom</wp:align>');
	expect(xml).toContain('descr="Arrangement sample"');
	expect(xml).toMatch(/<a:xfrm[^>]*rot="5400000"/);
});

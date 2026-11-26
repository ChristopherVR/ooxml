import { expect, type Locator, type Page } from '@playwright/test';
import { readFile } from 'node:fs/promises';
import { fileInput } from './helpers';

export async function choosePosition(editor: Locator, value: string) {
	await unfoldArrange(editor, 'Layout');
	await editor.getByRole('combobox', { name: 'Position', exact: true }).press('Enter');
	await editor.getByRole('menuitemradio', { name: new RegExp(value) }).click();
}

export async function openShapes(
	page: Page,
	fixture = new URL('./support/shapes.docx', import.meta.url),
) {
	await page.setViewportSize({ width: 2600, height: 950 });
	await page.goto('/?framework=vanilla');
	await (
		await fileInput(page)
	).setInputFiles({
		name: 'shapes.docx',
		mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
		buffer: await readFile(fixture),
	});
	const editor = page.locator('docx-editor');
	await expect(editor.locator('.ProseMirror .dve-shape').first()).toBeAttached();
	return editor;
}

export async function unfoldArrange(editor: Locator, tab: string) {
	const target = editor.getByRole('tab', { name: tab, exact: true });
	if ((await target.getAttribute('aria-selected')) !== 'true') await target.click();
	const panel = editor.getByRole('tabpanel', { name: tab, exact: true });
	const collapsed = panel.getByRole('button', { name: 'Arrange', exact: true });
	const alt = editor.getByRole('button', { name: 'Alt Text', exact: true });
	await panel.evaluate(
		() => new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve))),
	);
	if ((await collapsed.isVisible()) && (await collapsed.getAttribute('aria-expanded')) !== 'true')
		await collapsed.click();
	await expect(alt).toBeVisible();
}

import { expect, test, type Locator, type Page } from '@playwright/test';
import { readFile } from 'node:fs/promises';
import { dialogByHeading, fileInput } from './helpers';
import { savedDocument } from './saved-document';

test('Word arranges a floating shape from its Layout ribbon and saves the edit', async ({
	page,
}) => {
	const editor = await openShapes(page);
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

test('page geometry commands explain why they are unavailable in Draft', async ({ page }) => {
	const editor = await openShapes(page);
	await editor.locator('.ProseMirror .dve-shape').first().click();
	await editor.getByRole('tab', { name: 'Layout', exact: true }).click();
	const align = editor.getByRole('combobox', { name: 'Align', exact: true });
	await unfoldArrange(editor, 'Layout');
	await expect(align).toBeDisabled();
	await expect(align).toHaveAccessibleDescription('Switch to Print Layout to arrange drawings.');
});

test('ungroup keeps the native children on the canvas in consecutive stacking order', async ({
	page,
}) => {
	const fixture = new URL(
		'../../src/core/docx/__fixtures__/shapes/arrange-group.docx',
		import.meta.url,
	);
	const editor = await openShapes(page, fixture);
	await editor.getByRole('button', { name: 'Print Layout', exact: true }).click();
	const floats = editor.locator('.dve-print-float');
	await expect(floats).toHaveCount(1);
	await editor.getByRole('tab', { name: 'Layout', exact: true }).click();
	await unfoldArrange(editor, 'Layout');
	await editor.getByRole('button', { name: 'Selection Pane', exact: true }).click();
	const pane = editor.getByRole('complementary', { name: 'Selection Pane' });
	await pane.getByRole('listitem').getByRole('button').click();
	await unfoldArrange(editor, 'Layout');
	await editor.getByRole('combobox', { name: 'Group', exact: true }).selectOption('ungroup');
	await expect(floats).toHaveCount(3);
	const { xml } = await savedDocument(page, 'ungrouped.docx');
	const heights = [...xml.matchAll(/relativeHeight="(\d+)"/g)].map((match) => Number(match[1]));
	expect(heights).toHaveLength(3);
	expect(heights.slice(1)).toEqual(heights.slice(0, -1).map((height) => height + 1));
});

async function openShapes(page: Page, fixture = new URL('./support/shapes.docx', import.meta.url)) {
	await page.setViewportSize({ width: 1800, height: 950 });
	await page.goto('/?framework=vanilla');
	await (
		await fileInput(page)
	).setInputFiles({
		name: 'shapes.docx',
		mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
		buffer: await readFile(fixture),
	});
	const editor = page.locator('docx-editor');
	return editor;
}

async function unfoldArrange(editor: Locator, tab: string) {
	const panel = editor.getByRole('tabpanel', { name: tab, exact: true });
	const collapsed = panel.getByRole('button', { name: 'Arrange', exact: true });
	if (await collapsed.isVisible()) {
		const alt = editor.getByRole('button', { name: 'Alt Text', exact: true });
		if (!(await alt.isVisible())) await collapsed.click();
	}
}

import { expect, test, type Locator, type Page } from '@playwright/test';
import { readFile } from 'node:fs/promises';
import { dialogByHeading, fileInput } from './helpers';
import { savedDocument } from './saved-document';

test('Word arranges a floating shape from its Layout ribbon and saves the edit', async ({
	page,
}) => {
	const editor = await openShapes(page);
	await editor.getByRole('button', { name: 'Print Layout', exact: true }).click();
	const shape = editor.locator('.dve-print-float').first();
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

test('ordinary text edits leave unchanged Selection Pane rows mounted', async ({ page }) => {
	const editor = await openShapes(page);
	await editor.getByRole('tab', { name: 'Layout', exact: true }).click();
	await unfoldArrange(editor, 'Layout');
	await editor.getByRole('button', { name: 'Selection Pane', exact: true }).click();
	const row = editor
		.getByRole('complementary', { name: 'Selection Pane' })
		.getByRole('listitem')
		.first();
	await row.evaluate((element) => element.setAttribute('data-test-original-row', ''));
	const story = editor.locator('.ProseMirror').first();
	await story.press('Control+End');
	await story.press('x');
	await expect(story).toContainText('x');
	await expect
		.poll(() => row.evaluate((element) => element.hasAttribute('data-test-original-row')))
		.toBe(true);
});

test('Arrange controls fit the ribbon or its overflow in Shape Format and Layout', async ({
	page,
}) => {
	const editor = await openShapes(page);
	await editor.locator('.ProseMirror .dve-shape').first().click();
	for (const [tab, width] of [
		['Shape Format', 1800],
		['Layout', 2600],
		['Layout', 850],
	] as const) {
		await page.setViewportSize({ width, height: 950 });
		await editor.getByRole('tab', { name: tab, exact: true }).click();
		await unfoldArrange(editor, tab);
		const group = editor.locator('.ribbon-group[data-label="Arrange"]:visible').filter({
			has: editor.getByRole('combobox', { name: 'Position', exact: true }),
		});
		await expect(group.locator(':is(button, select)[data-arrange]:visible')).toHaveCount(9);
		await expect
			.poll(() =>
				group.evaluate((element) => {
					const ribbon = element.closest('.ribbon-panel, .ribbon-overflow-panel')!;
					const bounds = ribbon.getBoundingClientRect();
					return [...element.querySelectorAll<HTMLElement>(':is(button, select)[data-arrange]')]
						.filter((control) => control.getClientRects().length)
						.filter((control) => {
							const { left, right, top, bottom } = control.getBoundingClientRect();
							return (
								left < bounds.left - 1 ||
								right > bounds.right + 1 ||
								top < bounds.top - 1 ||
								bottom > bounds.bottom + 1
							);
						})
						.map((control) => control.getAttribute('aria-label'));
				}),
			)
			.toEqual([]);
	}
});

test('Position presets place an ellipse at the top, center and bottom of the printed page', async ({
	page,
}) => {
	const editor = await openShapes(page);
	await editor.getByRole('button', { name: 'Print Layout', exact: true }).click();
	const ellipse = editor.locator('.dve-print-float').nth(2);
	await expect(ellipse).toBeVisible();
	await ellipse.click({ force: true });
	await editor.getByRole('tab', { name: 'Layout', exact: true }).click();
	const measure = () =>
		ellipse.evaluate((element) => {
			const bounds = element.getBoundingClientRect();
			const sheet = element.closest('.dve-print-page')!;
			const margins = sheet.querySelector('.dve-print-column')!.getBoundingClientRect();
			return {
				top: bounds.top - margins.top,
				right: margins.right - bounds.right,
				bottom: margins.bottom - bounds.bottom,
				centerX: bounds.left + bounds.width / 2 - margins.left - margins.width / 2,
				centerY: bounds.top + bounds.height / 2 - margins.top - margins.height / 2,
			};
		});
	for (const [preset, edges] of [
		['Top Right', ['top', 'right']],
		['Bottom Right', ['bottom', 'right']],
		['Middle Center', ['centerX', 'centerY']],
	] as const) {
		await choosePosition(editor, preset);
		for (const edge of edges)
			await expect.poll(async () => Math.abs((await measure())[edge])).toBeLessThan(1);
	}
});

async function choosePosition(editor: Locator, value: string) {
	await unfoldArrange(editor, 'Layout');
	await editor.getByRole('combobox', { name: 'Position', exact: true }).press('Enter');
	await editor.getByRole('menuitemradio', { name: new RegExp(value) }).click();
}

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

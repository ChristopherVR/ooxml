import { expect, test } from '@playwright/test';
import { openShapes, unfoldArrange } from './word-arrange-support';
import { savedDocument } from './saved-document';

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
	await story.press('End');
	await story.press('Space');
	await story.press('a');
	await expect(async () => {
		const button = row.getByRole('button');
		await button.click();
		await expect(button).toHaveAttribute('aria-pressed', 'true', { timeout: 500 });
	}).toPass();
	await expect
		.poll(() => row.evaluate((element) => element.hasAttribute('data-test-original-row')))
		.toBe(true);
});

test('Selection Pane is docked and renames a drawing with F2 or a double-click', async ({
	page,
}) => {
	const editor = await openShapes(page);
	await editor.getByRole('tab', { name: 'Layout', exact: true }).click();
	await unfoldArrange(editor, 'Layout');
	await editor.getByRole('button', { name: 'Selection Pane', exact: true }).click();
	const pane = editor.getByRole('complementary', { name: 'Selection Pane' });
	const row = pane.getByRole('button', { name: 'Oval 3' });
	const dock = await pane.evaluate((element) => {
		const canvas = element.parentElement!.querySelector('.dve-canvas')!.getBoundingClientRect();
		const side = element.getBoundingClientRect();
		return [
			Math.abs(side.left - canvas.right),
			Math.abs(side.top - canvas.top),
			Math.abs(side.bottom - canvas.bottom),
		];
	});
	expect(Math.max(...dock)).toBeLessThan(1);
	await expect(pane.getByRole('heading', { name: 'Selection Pane' })).toBeVisible();
	await expect(pane.getByRole('button', { name: 'Close', exact: true })).toBeVisible();
	await editor.locator('office-ui-ribbon-actions [part="comments"]').click();
	const comments = editor.getByRole('complementary', { name: 'Comments' });
	await expect(comments).toBeVisible();
	await expect
		.poll(async () =>
			Math.abs(
				(await pane.boundingBox())!.x -
					((await comments.boundingBox())!.x + (await comments.boundingBox())!.width),
			),
		)
		.toBeLessThan(1);
	await row.dblclick();
	const name = pane.getByRole('textbox', { name: 'Drawing name' });
	await name.fill('Diagram oval');
	await name.press('Enter');
	const renamed = pane.getByRole('button', { name: 'Diagram oval' });
	await expect(renamed).toBeFocused();
	await renamed.press('F2');
	await name.fill('Final oval');
	await name.press('Enter');
	const selected = pane.getByRole('button', { name: 'Final oval', pressed: true });
	const colors = () =>
		selected.evaluate((element) => {
			const style = getComputedStyle(element);
			return [style.color, style.backgroundColor];
		});
	await editor.evaluate((element) => Object.assign(element, { theme: 'light' }));
	const light = await colors();
	expect(light[0]).not.toBe(light[1]);
	await editor.evaluate((element) => Object.assign(element, { theme: 'dark' }));
	await expect.poll(colors).not.toEqual(light);
	const dark = await colors();
	expect(dark[0]).not.toBe(dark[1]);
	await page.emulateMedia({ forcedColors: 'active' });
	await expect(selected).toHaveCSS('forced-color-adjust', 'none');
	const { xml } = await savedDocument(page, 'renamed-drawing.docx');
	expect(xml).toContain('name="Final oval"');
	expect(xml).not.toContain('title="Final oval"');
});

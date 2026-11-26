import { expect, test, type Page } from '@playwright/test';
import type { VisioViewerElement } from 'ooxml-ui/visio';
import { openDemo } from './demo-page';

async function drawingWithShape(page: Page) {
	await page.addInitScript(() => localStorage.clear());
	await page.setViewportSize({ width: 1440, height: 950 });
	await openDemo(page);
	const viewer = page.locator('visio-viewer');
	await viewer.evaluate(async (node) => {
		const editor = node as VisioViewerElement;
		await editor.createBlankDrawing();
		await editor.controller.applyEdits([
			{ type: 'create-rectangle', pageId: '0', shapeId: '1', x: 3, y: 4, width: 2, height: 1 },
		]);
	});
	await viewer.locator('svg.paper [data-shape-id="1"]').click();
	const dialog = viewer.locator('office-ui-dialog.stencil-dialog');
	const more = viewer.locator('#shapes-more');
	const menu = (scope: string, command: string) =>
		viewer.locator(`[data-shapes-menu="${scope}"] [command="${command}"]`);
	const confirm = async (label: string, value: string) => {
		await dialog.getByRole('textbox', { name: label }).fill(value);
		await dialog.getByRole('button', { name: 'OK', exact: true }).click();
	};
	return { viewer, more, menu, dialog, confirm };
}

test('returns focus to the invoking control after stencil dialogs and save', async ({ page }) => {
	test.setTimeout(60_000);
	const { viewer, more, menu, dialog, confirm } = await drawingWithShape(page);
	await more.press('Enter');
	await menu('more', 'new-stencil').click();
	await dialog.getByRole('textbox', { name: 'Stencil name' }).press('Escape');
	await expect(more).toBeFocused();
	await more.press('Enter');
	await menu('more', 'new-stencil').click();
	await confirm('Stencil name', 'Valves');
	const section = viewer.getByRole('region', { name: 'Valves' });
	const actions = section.getByRole('button', { name: 'Stencil actions' });
	await expect(section).toBeVisible();
	await expect(more).toBeFocused();
	await actions.press('Enter');
	await menu('stencil', 'add-selection').click();
	await confirm('Master name', 'Valve');
	const master = section.locator('[data-master]');
	await expect(master).toHaveText('Valve');
	await expect(actions).toBeFocused();
	await master.press('Shift+F10');
	await menu('master', 'rename').click();
	await confirm('Master name', 'Gate valve');
	await expect(master).toHaveText('Gate valve');
	await expect(master).toBeFocused();
	await master.press('Shift+F10');
	await menu('master', 'edit').click();
	await confirm('Default shape text', 'Open');
	await expect(master).toBeFocused();
	await actions.press('Enter');
	const download = page.waitForEvent('download');
	await menu('stencil', 'save-stencil-file').click();
	await download;
	await expect(actions).toBeFocused();
});

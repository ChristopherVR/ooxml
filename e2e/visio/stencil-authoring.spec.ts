import { readFile } from 'node:fs/promises';
import { expect, test, type Locator } from '@playwright/test';
import type { VisioViewerElement } from 'ooxml-ui/visio';
import { openDemo } from './demo-page';

test('creates, edits, saves and reopens a stencil through the Shapes pane', async ({ page }) => {
	test.setTimeout(60_000);
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

	const menu = (scope: string, command: string) =>
		viewer.locator(`[data-shapes-menu="${scope}"] [command="${command}"]`);
	const confirm = async (label: string, value: string) => {
		const dialog = viewer.locator('office-ui-dialog.stencil-dialog');
		await dialog.getByRole('textbox', { name: label }).fill(value);
		await dialog.getByRole('button', { name: 'OK', exact: true }).click();
	};
	await viewer.locator('#shapes-more').click();
	await menu('more', 'new-stencil').click();
	await confirm('Stencil name', 'Valves and fittings');
	const section = viewer.getByRole('region', { name: 'Valves and fittings', exact: true });
	const opened = viewer.locator('[data-stencil="file-1"]');
	const actions = (stencil: Locator) => stencil.getByRole('button', { name: 'Stencil actions' });
	await expect(section).toBeVisible();
	await actions(opened).press('Enter');
	await menu('stencil', 'add-selection').click();
	await confirm('Master name', 'Valve');
	const master = opened.locator('[data-master]');
	await expect(master).toHaveText('Valve');

	await master.press('Shift+F10');
	await menu('master', 'rename').click();
	await confirm('Master name', 'Spare valve');
	await expect(master).toHaveText('Spare valve');
	await master.press('Shift+F10');
	await menu('master', 'edit').click();
	await confirm('Default shape text', 'Normally closed');

	await actions(opened).press('Enter');
	const saved = page.waitForEvent('download');
	await menu('stencil', 'save-stencil-file').click();
	const download = await saved;
	expect(download.suggestedFilename()).toBe('Valves and fittings.vssx');
	const path = await download.path();
	await opened.locator('[data-close-stencil]').click();
	await viewer.locator('input[data-open-stencil]').setInputFiles({
		name: download.suggestedFilename(),
		mimeType: 'application/vnd.ms-visio.stencil',
		buffer: await readFile(path!),
	});

	const reopened = viewer.locator('[data-stencil="file-2"]');
	const valve = reopened.locator('[data-master]');
	await expect(valve).toHaveText('Spare valve');
	await valve.press('Enter');
	await expect
		.poll(() =>
			viewer.evaluate((node) => {
				const shapes = (node as VisioViewerElement).controller.state.document?.pages[0]?.shapes;
				const instance = shapes?.find((shape) => shape.masterId);
				return instance?.text.plainText;
			}),
		)
		.toBe('Normally closed');
	await actions(reopened).press('Enter');
	await menu('stencil', 'edit-stencil').click();
	await valve.press('Shift+F10');
	await menu('master', 'delete').click();
	await viewer
		.locator('office-ui-dialog.stencil-dialog')
		.getByRole('button', { name: 'Delete Master' })
		.click();
	await expect(valve).toHaveCount(0);
});

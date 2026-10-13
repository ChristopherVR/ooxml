import { expect, test } from '@playwright/test';
import { openDemo } from './demo-page';
import { downloadCopy } from './ribbon';

test('page context setup shares cancel, edit history and exported reopen', async ({ page }) => {
	await page.setViewportSize({ width: 1500, height: 950 });
	await openDemo(page);
	const viewer = page.locator('visio-viewer');
	await viewer.locator('office-ui-ribbon .file').click();
	await viewer.locator('[data-backstage-item="new"]').click();
	await viewer.locator('[data-backstage-action="new-blank"]').click();
	const paper = viewer.locator('svg.paper');
	await expect(paper).toBeVisible();
	const dialog = viewer.locator('.page-setup-dialog');
	const tabs = viewer.locator('.page-tabs');
	const open = async () => {
		await paper.click({ button: 'right', position: { x: 30, y: 30 } });
		const menu = viewer.locator('[data-context-menu="page"]');
		await expect(menu).toBeVisible();
		const item = menu.locator('[command="ctx-page-setup"]');
		await expect(item).not.toHaveAttribute('data-unsupported', '');
		await item.click();
		await expect(dialog).toBeVisible();
		await expect(menu).toBeHidden();
	};
	await open();
	await expect(dialog.getByRole('tab', { name: 'Print Setup', exact: true })).toHaveAttribute(
		'aria-selected',
		'true',
	);
	await dialog.getByRole('tab', { name: 'Page Properties' }).click();
	await dialog.getByRole('textbox', { name: 'Name', exact: true }).fill('Cancelled name');
	await dialog.getByRole('button', { name: 'Cancel', exact: true }).click();
	await expect(dialog).toBeHidden();
	await expect(tabs.getByRole('tab').first()).toHaveText('Page-1');
	await open();
	await dialog.getByRole('tab', { name: 'Page Properties' }).click();
	await dialog.getByRole('textbox', { name: 'Name', exact: true }).fill('Context plan');
	await dialog.getByRole('button', { name: 'OK', exact: true }).click();
	await expect(dialog).toBeHidden();
	await expect(tabs.getByRole('tab').first()).toHaveText('Context plan');
	await viewer.locator('.qat [data-command="undo"]').click();
	await expect(tabs.getByRole('tab').first()).toHaveText('Page-1');
	await viewer.locator('.qat [data-command="redo"]').click();
	await expect(tabs.getByRole('tab').first()).toHaveText('Context plan');
	const oldPaper = await paper.elementHandle();
	const pending = page.waitForEvent('download');
	await (await downloadCopy(viewer)).click();
	const stream = await (await pending).createReadStream();
	const chunks: Buffer[] = [];
	for await (const chunk of stream!) chunks.push(Buffer.from(chunk));
	await page.locator('#file').setInputFiles({
		name: 'context-plan.vsdx',
		mimeType: 'application/vnd.ms-visio.drawing',
		buffer: Buffer.concat(chunks),
	});
	await expect(page.locator('#file-name')).toHaveText('context-plan.vsdx');
	await expect.poll(() => oldPaper!.evaluate((node) => node.isConnected)).toBe(false);
	await expect(tabs.getByRole('tab').first()).toHaveText('Context plan');
	await open();
	await dialog.getByRole('tab', { name: 'Page Properties' }).click();
	await expect(dialog.getByRole('textbox', { name: 'Name', exact: true })).toHaveValue(
		'Context plan',
	);
	await dialog.getByRole('button', { name: 'Cancel', exact: true }).click();
});

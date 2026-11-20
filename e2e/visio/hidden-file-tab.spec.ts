import { expect, test, type Locator } from '@playwright/test';
import { openDemo } from './demo-page';

/** Sets the element's `customization`, the object every binding forwards as one prop. */
const hideFile = (viewer: Locator, hidden: boolean) =>
	viewer.evaluate((node, hide) => {
		(node as HTMLElement & { customization: unknown }).customization = hide
			? { ribbon: { hiddenTabs: ['file'] } }
			: null;
	}, hidden);
const backstageOpen = (viewer: Locator) =>
	viewer.evaluate(
		(node) =>
			(node.shadowRoot!.querySelector('office-ui-backstage') as HTMLElement & { open: boolean })
				.open,
	);

test.describe('customization.ribbon.hiddenTabs: file', () => {
	test('removes File on desktop and narrow widths; no key tip opens it', async ({ page }) => {
		await page.setViewportSize({ width: 1440, height: 900 });
		await openDemo(page);
		const viewer = page.locator('visio-viewer');
		await hideFile(viewer, true);
		await expect(viewer.locator('office-ui-ribbon .file')).toBeHidden();
		await expect(viewer.getByRole('tab', { name: 'Home', exact: true })).toBeVisible();
		await expect(viewer.getByRole('tab', { name: 'Design', exact: true })).toBeVisible();
		await viewer.locator('svg.paper').click({ position: { x: 5, y: 5 } });
		await page.keyboard.press('Alt');
		await page.keyboard.press('f');
		expect(await backstageOpen(viewer)).toBe(false);
		await page.setViewportSize({ width: 390, height: 844 });
		await expect(viewer.locator('office-ui-ribbon .file')).toBeHidden();
		await expect(viewer.getByRole('tab', { name: 'Home', exact: true })).toBeVisible();
	});

	test('Quick Access Save still announces Save to the host', async ({ page }) => {
		await openDemo(page);
		const viewer = page.locator('visio-viewer');
		await hideFile(viewer, true);
		await viewer.evaluate((node) => {
			(window as unknown as { commands: string[] }).commands = [];
			node.addEventListener('file-command', (event) => {
				event.preventDefault();
				(window as unknown as { commands: string[] }).commands.push(
					(event as CustomEvent<{ command: string }>).detail.command,
				);
			});
		});
		await viewer.locator('office-ui-title-bar').getByRole('button', { name: 'Save' }).click();
		await expect
			.poll(() => page.evaluate(() => (window as unknown as { commands: string[] }).commands))
			.toEqual(['save']);
		expect(await backstageOpen(viewer)).toBe(false);
	});

	test('hiding File while the backstage is open closes it and focuses the ribbon', async ({
		page,
	}) => {
		await openDemo(page);
		const viewer = page.locator('visio-viewer');
		await viewer.locator('office-ui-ribbon .file').click();
		await expect(viewer.locator('.backstage')).toBeVisible();
		await hideFile(viewer, true);
		await expect(viewer.locator('.backstage')).toBeHidden();
		await expect(viewer.getByRole('tab', { name: 'Home', exact: true })).toBeFocused();
		await hideFile(viewer, false);
		await expect(viewer.locator('office-ui-ribbon .file')).toBeVisible();
	});
});

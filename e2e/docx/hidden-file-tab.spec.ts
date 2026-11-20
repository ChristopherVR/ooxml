import { expect, test, type Locator } from '@playwright/test';
import { openSample } from './helpers';

/** Sets the element's `customization`, the object every binding forwards as one prop. */
const hideFile = (editor: Locator, hidden: boolean) =>
	editor.evaluate((element, hide) => {
		(element as HTMLElement & { customization: unknown }).customization = hide
			? { ribbon: { hiddenTabs: ['file'] } }
			: null;
	}, hidden);
const backstageOpen = (editor: Locator) =>
	editor.evaluate(
		(element) =>
			(element.shadowRoot!.querySelector('office-ui-backstage') as HTMLElement & { open: boolean })
				.open,
	);

test.describe('customization.ribbon.hiddenTabs: file', () => {
	test('removes File on desktop and narrow widths; no key tip or shortcut opens it', async ({
		page,
	}) => {
		await page.setViewportSize({ width: 1400, height: 800 });
		await openSample(page);
		const editor = page.locator('docx-editor');
		await hideFile(editor, true);
		await expect(editor.locator('.dve-ribbon .file')).toBeHidden();
		await expect(editor.getByRole('tab', { name: 'Home' })).toBeVisible();
		await expect(editor.getByRole('tab', { name: 'Insert' })).toBeVisible();
		await editor.locator('.ProseMirror').first().click();
		await page.keyboard.press('Alt');
		await page.keyboard.press('f');
		expect(await backstageOpen(editor)).toBe(false);
		await page.setViewportSize({ width: 390, height: 844 });
		await expect(editor.locator('.dve-ribbon .file')).toBeHidden();
		await expect(editor.getByRole('tab', { name: 'Home' })).toBeVisible();
	});

	test('Ctrl+S still announces Save to the host', async ({ page }) => {
		await openSample(page);
		const editor = page.locator('docx-editor');
		await hideFile(editor, true);
		await editor.evaluate((element) => {
			(window as unknown as { commands: string[] }).commands = [];
			element.addEventListener('file-command', (event) => {
				event.preventDefault();
				(window as unknown as { commands: string[] }).commands.push(
					(event as CustomEvent<{ command: string }>).detail.command,
				);
			});
		});
		await editor.locator('.ProseMirror').first().click();
		await page.keyboard.press('ControlOrMeta+s');
		await expect
			.poll(() => page.evaluate(() => (window as unknown as { commands: string[] }).commands))
			.toEqual(['save']);
	});

	test('hiding File while the backstage is open closes it and focuses the ribbon', async ({
		page,
	}) => {
		await openSample(page);
		const editor = page.locator('docx-editor');
		await editor.locator('.dve-ribbon .file').click();
		await expect(editor.locator('.dve-backstage')).toBeVisible();
		await hideFile(editor, true);
		await expect(editor.locator('.dve-backstage')).toBeHidden();
		await expect(editor.getByRole('tab', { name: 'Home' })).toBeFocused();
		await hideFile(editor, false);
		await expect(editor.locator('.dve-ribbon .file')).toBeVisible();
	});
});

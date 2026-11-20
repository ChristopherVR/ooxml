import { readFileSync } from 'node:fs';
import { expect, test, type Locator } from '@playwright/test';
import { editor, grid, landingFileInput, openLanding, openSample, part } from './helpers';

// An Excel-signed sample (Apache POI test data): its signature verifies, its signer is not trusted.
const SIGNED = readFileSync(
	new URL(
		'../../src/core/opc/signature/__fixtures__/office/hello-world-signed.xlsx',
		import.meta.url,
	),
);

/** Sets the element's `customization`, the object every binding forwards as one prop. */
const hideFile = (element: Locator, hidden: boolean) =>
	element.evaluate((node, hide) => {
		(node as HTMLElement & { customization: unknown }).customization = hide
			? { ribbon: { hiddenTabs: ['file'] } }
			: null;
	}, hidden);
const backstageOpen = (element: Locator) =>
	element.evaluate(
		(node) =>
			(node.shadowRoot!.querySelector('office-ui-backstage') as HTMLElement & { open: boolean })
				.open,
	);
const fileButton = (element: Locator) => element.getByRole('button', { name: 'File', exact: true });

test.describe('customization.ribbon.hiddenTabs: file', () => {
	test('removes File on desktop and narrow widths; no key tip or shortcut opens it', async ({
		page,
	}) => {
		await page.setViewportSize({ width: 1400, height: 800 });
		await openSample(page);
		const xlsx = editor(page);
		await hideFile(xlsx, true);
		await expect(fileButton(xlsx)).toBeHidden();
		await expect(xlsx.getByRole('tab', { name: 'Home' })).toBeVisible();
		await expect(xlsx.getByRole('tab', { name: 'Formulas' })).toBeVisible();
		await grid(page).click();
		await page.keyboard.press('Alt');
		await page.keyboard.press('f');
		expect(await backstageOpen(xlsx)).toBe(false);
		await page.setViewportSize({ width: 390, height: 844 });
		await expect(fileButton(xlsx)).toBeHidden();
		await expect(xlsx.getByRole('tab', { name: 'Home' })).toBeVisible();
		await expect(part(page, 'ribbon')).toBeVisible();
	});

	test('Ctrl+S and Ctrl+P still reach the host', async ({ page }) => {
		await openSample(page);
		const xlsx = editor(page);
		await hideFile(xlsx, true);
		await xlsx.evaluate((node) => {
			(window as unknown as { commands: string[] }).commands = [];
			node.addEventListener('file-command', (event) => {
				event.preventDefault();
				(window as unknown as { commands: string[] }).commands.push(
					(event as CustomEvent<{ command: string }>).detail.command,
				);
			});
		});
		await grid(page).click();
		await page.keyboard.press('ControlOrMeta+s');
		await page.keyboard.press('ControlOrMeta+p');
		await expect
			.poll(() => page.evaluate(() => (window as unknown as { commands: string[] }).commands))
			.toEqual(['save', 'print']);
		expect(await backstageOpen(xlsx)).toBe(false);
	});

	test('hiding File while the backstage is open closes it and focuses the ribbon', async ({
		page,
	}) => {
		await openSample(page);
		const xlsx = editor(page);
		await fileButton(xlsx).click();
		await expect.poll(() => backstageOpen(xlsx)).toBe(true);
		await hideFile(xlsx, true);
		expect(await backstageOpen(xlsx)).toBe(false);
		await expect(xlsx.getByRole('tab', { name: 'Home' })).toBeFocused();
		await hideFile(xlsx, false);
		await expect(fileButton(xlsx)).toBeVisible();
	});

	test('keeps the signature state on screen, without actions', async ({ page }) => {
		await openLanding(page);
		await (
			await landingFileInput(page)
		).setInputFiles({
			name: 'signed.xlsx',
			mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
			buffer: SIGNED,
		});
		await expect(grid(page)).toBeVisible();
		const xlsx = editor(page);
		const badge = xlsx.locator('[data-testid="signature-badge"]');
		const notice = xlsx.locator('[data-testid="signature-notice"]');
		await expect(badge.getByRole('button')).toBeVisible();
		await hideFile(xlsx, true);
		await expect(badge).toContainText('Signed, signer not verified');
		await expect(badge.getByRole('button')).toHaveCount(0);
		await expect(notice).toBeVisible();
		await expect(notice.getByRole('button', { name: 'View signatures' })).toHaveCount(0);
		await hideFile(xlsx, false);
		await expect(notice.getByRole('button', { name: 'View signatures' })).toBeVisible();
	});
});

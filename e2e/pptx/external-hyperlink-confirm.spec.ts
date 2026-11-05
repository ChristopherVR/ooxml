/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright spec */
/**
 * Trust Center > "Confirm before opening external hyperlinks": following a linked shape asks in
 * an in-app dialog (never the browser's confirm), in every binding. Cancel opens nothing; Open
 * opens the address in a new tab; Escape cancels.
 */
import { fileURLToPath } from 'node:url';

import { expect, test } from '@playwright/test';
import type { Page } from '@playwright/test';
import type { PowerPointViewerAPI } from 'ooxml-ui/pptx';

import { resetTabSession, slideStage } from './support/deck';

declare global {
	interface Window {
		/** Public component handle exposed by each development demo. */
		__pptxViewer: PowerPointViewerAPI;
	}
}

// One slide; its rectangle links to https://example.com/hyperlink-fixture.
const fixture = fileURLToPath(new URL('./fixtures/hyperlink-action.pptx', import.meta.url));

async function open(page: Page): Promise<void> {
	// A browser dialog would block the page, so any that appears fails the test.
	page.on('dialog', (dialog) => {
		void dialog.dismiss();
		throw new Error(`browser ${dialog.type()} dialog: ${dialog.message()}`);
	});
	await resetTabSession(page);
	await page.goto('/');
	await page.locator('#file-input').setInputFiles(fixture);
	await expect(slideStage(page)).toBeVisible();
	await expect
		.poll(() =>
			page.evaluate(() => window.__pptxViewer.getElements(0).some((el) => el.actionClick)),
		)
		.toBe(true);
	await page.evaluate(() => window.__pptxViewer.setMode('preview'));
}

async function clickLinkedShape(page: Page): Promise<void> {
	const id = await page.evaluate(
		() => window.__pptxViewer.getElements(0).find((el) => el.actionClick)?.id ?? '',
	);
	await slideStage(page).locator(`[data-element-id="${id}"]`).first().click();
}

test('a linked shape asks in an in-app dialog and opens only after Open', async ({
	page,
	context,
}) => {
	await open(page);
	const opened: string[] = [];
	context.on('page', (created) => opened.push(created.url()));

	await clickLinkedShape(page);
	const dialog = page.getByRole('alertdialog');
	await expect(dialog).toBeVisible();
	await expect(dialog).toContainText('https://example.com/hyperlink-fixture');
	await expect(dialog).toHaveAttribute('aria-modal', 'true');

	// Cancel (and Escape) keep the viewer where it is and open nothing.
	await dialog.getByRole('button').first().click();
	await expect(dialog).toHaveCount(0);
	await clickLinkedShape(page);
	await expect(dialog).toBeVisible();
	await page.keyboard.press('Escape');
	await expect(dialog).toHaveCount(0);
	expect(opened).toHaveLength(0);

	// Open follows the link in a new tab.
	await clickLinkedShape(page);
	await expect(dialog).toBeVisible();
	const popup = context.waitForEvent('page');
	await dialog.getByRole('button').last().click();
	await expect(dialog).toHaveCount(0);
	const tab = await popup;
	expect(tab.url()).toMatch(/example\.com\/hyperlink-fixture|about:blank/u);
});

import { expect, test } from '@playwright/test';
import { readFile } from 'node:fs/promises';
import type { TeamsApp } from '../../src/ui/src/teams/app/teams-app';

async function seedChannels(page: import('@playwright/test').Page) {
	await page.goto(`/?local=1&name=Ada&room=compact-navigation-${Date.now()}`);
	await expect(page.getByRole('heading', { name: '# General', exact: true })).toBeVisible({
		timeout: 15_000,
	});
	await page.locator('teams-app').evaluate((element) => {
		const client = (element as TeamsApp).client!;
		const general = client.getState().selectedChannelId;
		client.createChannel('Design');
		client.select(general);
	});
	await page.setViewportSize({ width: 390, height: 844 });
}

test('navigates compact channels and calls through the list view with a bottom bar', async ({
	page,
}) => {
	await seedChannels(page);
	const back = page.getByRole('button', { name: 'Back to channels', exact: true });
	const bar = page.getByRole('navigation', { name: 'App bar', exact: true });
	const list = page.locator('teams-app office-ui-channel-list');
	const message = page.getByRole('textbox', { name: 'Message', exact: true });
	// The rail is a bottom bar, not a left column.
	const barBox = (await bar.boundingBox())!;
	expect(barBox.width).toBeGreaterThan(380);
	expect(barBox.y + barBox.height).toBeGreaterThan(830);
	await message.fill('Keep the General draft');
	await back.click();
	await expect(list).toBeVisible();
	await expect(message).toBeHidden();
	await page.screenshot({ path: test.info().outputPath('compact-channel-list.png') });
	await list.getByRole('button', { name: 'Design', exact: true }).click();
	await expect(page.getByRole('heading', { name: '# Design', exact: true })).toBeVisible();
	await expect(list).toBeHidden();
	await message.fill('Keep the Design draft');
	await back.click();
	await list.getByRole('button', { name: 'General', exact: true }).click();
	await expect(message).toHaveValue('Keep the General draft');
	await back.click();
	await list.getByRole('button', { name: 'Add channel', exact: true }).click();
	await page.getByRole('textbox', { name: 'Channel name', exact: true }).fill('Planning');
	await page.getByRole('button', { name: 'Create', exact: true }).click();
	await expect(page.getByRole('heading', { name: '# Planning', exact: true })).toBeFocused();
	await bar.getByRole('button', { name: 'More', exact: true }).click();
	await page
		.getByRole('dialog', { name: 'More apps' })
		.getByRole('button', { name: 'Calls', exact: true })
		.click();
	await expect(page.getByRole('button', { name: 'Meet now', exact: true })).toHaveCount(3);
	await bar.getByRole('button', { name: 'Teams', exact: true }).click();
	await expect(list).toBeVisible();
	await page.setViewportSize({ width: 1280, height: 844 });
	await expect(back).toBeHidden();
	await expect(page.getByRole('heading', { name: '# Planning', exact: true })).toBeVisible();
	await page.getByRole('button', { name: 'Design', exact: true }).click();
	await expect(message).toHaveValue('Keep the Design draft');
});

test('keeps an edited workbook mounted when compact navigation is cancelled', async ({ page }) => {
	test.setTimeout(90_000);
	const bytes = await readFile(
		new URL('../../src/core/xlsx/__fixtures__/openpyxl-features.xlsx', import.meta.url),
	);
	await page.route('https://files.test/Budget.xlsx', (route) =>
		route.fulfill({
			body: bytes,
			contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
		}),
	);
	await seedChannels(page);
	await page.locator('teams-app').evaluate((element) =>
		(element as TeamsApp).previewContent({
			attachment: { name: 'Budget.xlsx', kind: 'xlsx', url: 'https://files.test/Budget.xlsx' },
			url: 'https://files.test/Budget.xlsx',
		}),
	);
	await page.getByRole('button', { name: 'Edit workbook', exact: true }).click({ timeout: 60_000 });
	const formula = page.getByRole('textbox', { name: 'Formula Bar', exact: true });
	await formula.click();
	await formula.press('ControlOrMeta+A');
	await formula.pressSequentially('Keep this workbook edit');
	await formula.press('Enter');
	const nameBox = page.getByRole('combobox', { name: 'Name Box', exact: true });
	await nameBox.fill('A1');
	await nameBox.press('Enter');
	const teamsTab = page
		.getByRole('navigation', { name: 'App bar', exact: true })
		.getByRole('button', { name: 'Teams', exact: true });
	const list = page.locator('teams-app office-ui-channel-list');
	page.once('dialog', (dialog) => dialog.dismiss());
	await teamsTab.click();
	await expect(list).toBeHidden();
	await expect(formula).toHaveValue('Keep this workbook edit');
	await page.setViewportSize({ width: 1280, height: 844 });
	await expect(formula).toBeVisible();
	await expect(formula).toHaveValue('Keep this workbook edit');
	await page.setViewportSize({ width: 390, height: 844 });
	page.once('dialog', (dialog) => dialog.accept());
	await teamsTab.click();
	await expect(list).toBeVisible();
	await expect(formula).not.toBeVisible();
});

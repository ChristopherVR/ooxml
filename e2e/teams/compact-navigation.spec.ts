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

test('opens compact channel and call navigation with drafts, keyboard focus and resize handling', async ({
	page,
}) => {
	await seedChannels(page);
	const trigger = page.getByRole('button', { name: 'Open navigation', exact: true });
	const drawer = page.getByRole('dialog', { name: 'Workspace navigation', exact: true });
	const navigation = page.locator('teams-navigation-drawer');
	const message = page.getByRole('textbox', { name: 'Message', exact: true });
	await message.fill('Keep the General draft');
	await trigger.click();
	await expect(drawer).toBeVisible();
	await expect(trigger).toHaveAttribute('aria-expanded', 'true');
	const close = drawer.getByRole('button', { name: 'Close navigation', exact: true });
	await expect(close).toBeFocused();
	await navigation.getByRole('button', { name: 'Channels', exact: true }).click();
	await close.focus();
	await close.press('Shift+Tab');
	await expect(navigation.getByRole('button', { name: 'Add channel', exact: true })).toBeFocused();
	await navigation.getByRole('button', { name: 'Channels', exact: true }).click();
	await close.focus();
	await close.press('Shift+Tab');
	await expect(navigation.getByRole('button', { name: 'Design', exact: true })).toBeFocused();
	await page.keyboard.press('Escape');
	await expect(drawer).not.toBeVisible();
	await expect(trigger).toBeFocused();
	await expect(message).toHaveValue('Keep the General draft');
	await trigger.click();
	await page.mouse.click(385, 300);
	await expect(drawer).not.toBeVisible();
	await expect(trigger).toBeFocused();
	await trigger.click();
	await page.screenshot({ path: test.info().outputPath('compact-channel-navigation.png') });
	await navigation.getByRole('button', { name: 'Design', exact: true }).click();
	await expect(page.getByRole('heading', { name: '# Design', exact: true })).toBeFocused();
	await expect(drawer).not.toBeVisible();
	await message.fill('Keep the Design draft');
	await trigger.click();
	await navigation.getByRole('button', { name: 'General', exact: true }).click();
	await expect(message).toHaveValue('Keep the General draft');
	await trigger.click();
	await navigation.getByRole('button', { name: 'Add channel', exact: true }).click();
	await page.getByRole('textbox', { name: 'Channel name', exact: true }).fill('Planning');
	await page.getByRole('button', { name: 'Create', exact: true }).click();
	await expect(page.getByRole('heading', { name: '# Planning', exact: true })).toBeFocused();
	await expect(drawer).not.toBeVisible();
	await page
		.getByRole('navigation', { name: 'App bar', exact: true })
		.getByRole('button', { name: 'Calls', exact: true })
		.click();
	await trigger.click();
	await expect(navigation.getByRole('button', { name: 'Meet now', exact: true })).toHaveCount(3);
	await close.click();
	await expect(trigger).toBeFocused();
	await page
		.getByRole('navigation', { name: 'App bar', exact: true })
		.getByRole('button', { name: 'Teams', exact: true })
		.click();
	await trigger.click();
	await page.setViewportSize({ width: 1280, height: 844 });
	await expect(drawer).not.toBeVisible();
	await expect(trigger).toBeHidden();
	await expect(page.getByRole('heading', { name: '# Planning', exact: true })).toBeFocused();
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
	const trigger = page.getByRole('button', { name: 'Open navigation', exact: true });
	const drawer = page.getByRole('dialog', { name: 'Workspace navigation', exact: true });
	const navigation = page.locator('teams-navigation-drawer');
	await trigger.click();
	await navigation.getByRole('button', { name: 'Add channel', exact: true }).click();
	const creation = page.getByRole('dialog', { name: 'Create a channel', exact: true });
	await creation
		.getByRole('textbox', { name: 'Channel name', exact: true })
		.fill('Keep these details');
	await creation
		.getByRole('textbox', { name: 'Description (optional)', exact: true })
		.fill('Keep this description');
	page.once('dialog', (dialog) => dialog.dismiss());
	await creation.getByRole('button', { name: 'Create', exact: true }).click();
	await expect(creation).toBeVisible();
	await expect(creation.getByRole('textbox', { name: 'Channel name', exact: true })).toHaveValue(
		'Keep these details',
	);
	await expect(
		creation.getByRole('textbox', { name: 'Description (optional)', exact: true }),
	).toHaveValue('Keep this description');
	await expect(formula).toHaveValue('Keep this workbook edit');
	await creation.getByRole('button', { name: 'Cancel', exact: true }).click();
	page.once('dialog', (dialog) => dialog.dismiss());
	await navigation.getByRole('button', { name: 'Design', exact: true }).click();
	await expect(drawer).toBeVisible();
	await expect(formula).toHaveValue('Keep this workbook edit');
	await drawer.getByRole('button', { name: 'Close navigation', exact: true }).click();
	await expect(formula).toBeVisible();
	await expect(formula).toHaveValue('Keep this workbook edit');
	await trigger.click();
	await page.setViewportSize({ width: 1280, height: 844 });
	await expect(drawer).not.toBeVisible();
	await expect(page.getByRole('main')).toBeFocused();
	await expect(formula).toHaveValue('Keep this workbook edit');
	await page.setViewportSize({ width: 390, height: 844 });
	await trigger.click();
	page.once('dialog', (dialog) => dialog.accept());
	await navigation.getByRole('button', { name: 'Design', exact: true }).click();
	await expect(drawer).not.toBeVisible();
	await expect(page.getByRole('heading', { name: '# Design', exact: true })).toBeFocused();
	await expect(formula).not.toBeVisible();
});

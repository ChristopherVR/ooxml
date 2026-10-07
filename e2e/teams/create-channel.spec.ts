import { expect, test } from '@playwright/test';
import type { TeamsApp } from '../../src/ui/src/teams/app/teams-app';

test('creates standard channels through a dialog with descriptions and cancel restoration', async ({
	page,
}) => {
	await page.goto(`/?local=1&name=Ada&room=create-channel-${Date.now()}`);
	await expect(page.getByRole('heading', { name: '# General', exact: true })).toBeVisible({
		timeout: 15_000,
	});
	for (const width of [1280, 390]) {
		await page.setViewportSize({ width, height: 844 });
		const open = async () => {
			if (width === 390)
				await page.getByRole('button', { name: 'Open navigation', exact: true }).click();
			await page.getByRole('button', { name: 'Add channel', exact: true }).click();
		};
		await open();
		const dialog = page.getByRole('dialog', { name: 'Create a channel', exact: true });
		const name = dialog.getByRole('textbox', { name: 'Channel name', exact: true });
		await expect(name).toBeFocused();
		await dialog.getByRole('button', { name: 'Create', exact: true }).click();
		await expect(dialog.getByRole('alert')).toHaveText('Enter a channel name.');
		await name.fill('Cancelled');
		await dialog.getByRole('button', { name: 'Cancel', exact: true }).click();
		await expect(dialog).not.toBeVisible();
		await expect(page.getByRole('button', { name: 'Add channel', exact: true })).toBeFocused();
		// The navigation drawer stays open after cancelling its child dialog.
		if (width === 390)
			await page.getByRole('button', { name: 'Close navigation', exact: true }).click();
		await open();
		await expect(name).toHaveValue('');
		const channelName = `Planning ${width}`;
		const topic = 'Budget planning\n' + 'a'.repeat(300);
		await name.fill(channelName);
		await dialog.getByRole('textbox', { name: 'Description (optional)', exact: true }).fill(topic);
		await page.screenshot({ path: test.info().outputPath(`create-channel-${width}.png`) });
		await dialog.getByRole('button', { name: 'Create', exact: true }).click();
		await expect(dialog).not.toBeVisible();
		await expect(
			page.getByRole('heading', { name: `# ${channelName}`, exact: true }),
		).toBeFocused();
		await expect(
			page.getByRole('dialog', { name: 'Workspace navigation', exact: true }),
		).not.toBeVisible();
		expect(
			await page
				.locator('teams-app')
				.evaluate((element) => (element as TeamsApp).client!.getState().channel?.topic),
		).toBe(topic);
	}
	await page.reload();
	await expect(page.getByRole('heading', { name: '# Planning 390', exact: true })).toBeVisible({
		timeout: 15_000,
	});
	expect(
		await page
			.locator('teams-app')
			.evaluate((element) =>
				(element as TeamsApp)
					.client!.getState()
					.channels.some((channel) => channel.name === 'Cancelled'),
			),
	).toBe(false);
});

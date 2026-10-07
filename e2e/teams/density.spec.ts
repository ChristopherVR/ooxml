import { expect, test } from '@playwright/test';
import type { TeamsApp } from '../../src/ui/src/teams/app/teams-app';

test('saves personal chat density and changes real posts and thread spacing', async ({ page }) => {
	const room = `density-${Date.now()}`;
	await page.goto(`/?local=1&name=Ada&room=${room}`);
	await expect(page.getByRole('heading', { name: '# General', exact: true })).toBeVisible({
		timeout: 15_000,
	});
	await page.locator('teams-app').evaluate((element) => {
		const client = (element as TeamsApp).client!;
		const channel = client.getState().selectedChannelId;
		const root = client.workspace.chat.post(channel, { text: 'Density review' })!;
		client.workspace.chat.post(channel, { text: 'Thread detail', replyTo: root.id });
	});
	await page.getByRole('button', { name: '1 reply', exact: true }).click();
	const lists = page.locator('office-ui-chat-list');
	await expect(lists).toHaveCount(2);
	const sizes = () =>
		lists.evaluateAll((elements) =>
			elements.map((element) => {
				const bubble = element.shadowRoot!.querySelector<HTMLElement>('.bubble')!;
				const message = element.shadowRoot!.querySelector<HTMLElement>('.msg')!;
				return {
					height: bubble.getBoundingClientRect().height,
					gap: parseFloat(getComputedStyle(message).marginTop),
				};
			}),
		);
	const comfy = await sizes();
	await page.getByRole('button', { name: 'Settings and more', exact: true }).click();
	await page.getByRole('menuitem', { name: 'Settings', exact: true }).click();
	await page.getByRole('tab', { name: 'Appearance and accessibility', exact: true }).click();
	await expect(page.getByRole('radio', { name: /Comfy/ })).toBeChecked();
	await page.getByRole('radio', { name: /Comfy/ }).focus();
	await page.getByRole('radio', { name: /Comfy/ }).press('ArrowRight');
	await expect(page.getByRole('radio', { name: /Compact/ })).toBeChecked();
	await expect(page.locator('teams-app')).toHaveAttribute('data-chat-density', 'compact');
	await page.screenshot({ path: test.info().outputPath('settings-density.png') });
	await page.getByRole('button', { name: 'Close settings', exact: true }).click();
	const compact = await sizes();
	for (const [index, value] of compact.entries()) {
		expect(value.height).toBeLessThan(comfy[index]!.height);
		expect(value.gap).toBeLessThan(comfy[index]!.gap);
	}
	await page.screenshot({ path: test.info().outputPath('compact-thread.png') });
	await page.locator('teams-app').evaluate(async (element) => {
		await (element as TeamsApp).client!.send({ text: 'Keep this immediate reload' });
	});
	await page.reload();
	await expect(page.locator('teams-app')).toHaveAttribute('data-chat-density', 'compact');
	await expect(page.getByText('Keep this immediate reload', { exact: true })).toBeVisible();
	await page.getByRole('button', { name: 'Settings and more', exact: true }).click();
	await page.getByRole('menuitem', { name: 'Settings', exact: true }).click();
	await page.getByRole('tab', { name: 'Appearance and accessibility', exact: true }).click();
	await expect(page.getByRole('radio', { name: /Compact/ })).toBeChecked();
	await page.getByRole('button', { name: 'Close settings', exact: true }).click();
	const original = await page.locator('teams-app').evaluate((element) => {
		const app = element as TeamsApp;
		const id = app.userId;
		app.userId = 'different-density-user';
		return id;
	});
	await expect(page.locator('teams-app')).toHaveAttribute('data-chat-density', 'comfy');
	await page.locator('teams-app').evaluate((element, id) => {
		(element as TeamsApp).userId = id;
	}, original);
	await expect(page.locator('teams-app')).toHaveAttribute('data-chat-density', 'compact');
	await page.getByRole('button', { name: 'Settings and more', exact: true }).click();
	await page.getByRole('menuitem', { name: 'Settings', exact: true }).click();
	await page.getByRole('tab', { name: 'Appearance and accessibility', exact: true }).click();
	await page.getByRole('radio', { name: /Comfy/ }).check();
	await page.getByRole('button', { name: 'Close settings', exact: true }).click();
	await page.getByRole('button', { name: '1 reply', exact: true }).click();
	await expect(lists).toHaveCount(2);
	expect(await sizes()).toEqual(comfy);
});

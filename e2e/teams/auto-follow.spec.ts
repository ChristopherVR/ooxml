import { expect, test } from '@playwright/test';
import type { TeamsApp } from '../../src/ui/src/teams/app/teams-app';

test('automatically follows sent threads with persistent independent settings', async ({
	page,
}) => {
	const room = `auto-follow-${Date.now()}`;
	await page.goto(`/?local=1&name=Ada&room=${room}`);
	await expect(page.getByRole('heading', { name: '# General', exact: true })).toBeVisible({
		timeout: 15_000,
	});
	await page
		.getByRole('textbox', { name: 'Message', exact: true })
		.fill('Automatically followed post');
	await page.getByRole('button', { name: 'Send message', exact: true }).click();
	await page.getByRole('button', { name: 'Followed threads', exact: true }).click();
	const feed = page.getByRole('region', { name: 'Followed threads', exact: true });
	await expect(feed.getByText('Automatically followed post', { exact: true })).toBeVisible();
	await feed.getByRole('checkbox', { name: 'Threads I start', exact: true }).uncheck();
	await feed.getByRole('checkbox', { name: 'Threads I reply to', exact: true }).uncheck();
	await feed.getByRole('button', { name: 'Unfollow thread', exact: true }).click();
	await page.locator('teams-app').evaluate(async (element) => {
		const client = (element as TeamsApp).client!;
		await client.send({ text: 'Post with automatic following disabled' });
		client.startReply(client.getState().posts[0]!.id);
		await client.send({ text: 'Reply with automatic following disabled' });
	});
	await expect(feed.getByRole('button', { name: 'Unfollow thread', exact: true })).toHaveCount(0);
	await expect
		.poll(() => page.evaluate((room) => localStorage.getItem(`teams:doc:${room}`), room))
		.not.toBeNull();
	await page.reload();
	await page.getByRole('button', { name: 'Followed threads', exact: true }).click();
	await expect(
		feed.getByRole('checkbox', { name: 'Threads I start', exact: true }),
	).not.toBeChecked();
	await expect(
		feed.getByRole('checkbox', { name: 'Threads I reply to', exact: true }),
	).not.toBeChecked();
	await feed.getByRole('checkbox', { name: 'Threads I reply to', exact: true }).check();
	await page.locator('teams-app').evaluate(async (element) => {
		const client = (element as TeamsApp).client!;
		client.startReply(client.getState().posts[0]!.id);
		await client.send({ text: 'Reply with automatic following enabled' });
	});
	await expect(feed.getByText('Automatically followed post', { exact: true })).toBeVisible();
	await page.screenshot({ path: test.info().outputPath('followed-threads.png') });
});

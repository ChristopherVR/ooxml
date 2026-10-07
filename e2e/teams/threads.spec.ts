import { expect, test } from '@playwright/test';
import type { TeamsApp } from '../../src/ui/src/teams/app/teams-app';

test('opens channel threads, synchronizes replies and retains deleted parents', async ({
	page,
}) => {
	await page.goto(`/?local=1&name=Ada&room=threads-${Date.now()}`);
	await expect(page.getByRole('heading', { name: '# General', exact: true })).toBeVisible({
		timeout: 15_000,
	});
	const ids = await page.locator('teams-app').evaluate(async (element) => {
		const client = (element as TeamsApp).client!;
		client.createChannel('Threads');
		await new Promise((resolve) => setTimeout(resolve, 0));
		const channel = client.getState().selectedChannelId;
		const root = client.workspace.chat.post(channel, { text: 'Budget discussion' })!;
		const child = client.workspace.chat.post(channel, {
			text: 'Initial review',
			replyTo: root.id,
		})!;
		client.workspace.chat.post(channel, { text: 'Nested review', replyTo: child.id });
		return { channel, root: root.id };
	});
	await expect(page.getByText('Budget discussion', { exact: true })).toBeVisible();
	await expect(page.getByText('Initial review', { exact: true })).toHaveCount(0);
	await page.getByRole('button', { name: '2 replies', exact: true }).click();
	const thread = page.getByRole('complementary', { name: 'Thread', exact: true });
	await expect(thread.getByRole('heading', { name: 'Thread', exact: true })).toBeFocused();
	await expect(thread.getByText('Initial review', { exact: true })).toBeVisible();
	await expect(thread.getByText('Nested review', { exact: true })).toBeVisible();
	await thread.getByRole('textbox', { name: 'Message', exact: true }).fill('Thread response');
	await thread.getByRole('button', { name: 'Send message', exact: true }).click();
	await expect(thread.getByText('Thread response', { exact: true })).toBeVisible();
	const peer = await page.context().newPage();
	await peer.goto(page.url().replace('name=Ada', 'name=Bob'));
	await expect
		.poll(
			() =>
				peer
					.locator('teams-app')
					.evaluate((element) =>
						(element as TeamsApp).client
							?.getState()
							.channels.some((channel) => channel.name === 'Threads'),
					),
			{ timeout: 15_000 },
		)
		.toBe(true);
	await peer.locator('teams-app').evaluate((element, ids) => {
		const client = (element as TeamsApp).client!;
		client.select(ids.channel);
		client.workspace.chat.post(ids.channel, { text: 'Peer review', replyTo: ids.root });
	}, ids);
	await expect(thread.getByText('Peer review', { exact: true })).toBeVisible();
	await page.screenshot({ path: test.info().outputPath('channel-thread.png') });
	await page.setViewportSize({ width: 800, height: 800 });
	await expect(page.locator('.conversation-main')).toBeHidden();
	await expect(thread.getByRole('textbox', { name: 'Message', exact: true })).toBeVisible();
	expect(
		await peer
			.locator('teams-app')
			.evaluate((element) => (element as TeamsApp).client!.getState().thread),
	).toBeNull();
	await page
		.locator('teams-app')
		.evaluate((element, root) => (element as TeamsApp).client!.deleteMessage(root), ids.root);
	await expect(thread.getByText('This message was deleted', { exact: true }).first()).toBeVisible();
	await expect(thread.getByText('Peer review', { exact: true })).toBeVisible();
	await thread.getByRole('button', { name: 'Close thread', exact: true }).click();
	await expect(thread).toHaveCount(0);
	await expect(page.getByRole('button', { name: '4 replies', exact: true })).toBeVisible();
	await expect(page.getByRole('button', { name: '4 replies', exact: true })).toBeFocused();
	await page.getByRole('searchbox', { name: 'Search messages', exact: true }).fill('Peer review');
	await page.getByRole('listbox', { name: 'Search results' }).getByRole('button').click();
	await expect(thread.getByText('Peer review', { exact: true })).toBeVisible();
	await peer.close();
});

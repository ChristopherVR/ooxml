import { expect, test } from '@playwright/test';
import type { TeamsApp } from '../../src/ui/src/teams/app/teams-app';

test('preserves post and thread drafts, finds them centrally and requires restored files', async ({
	page,
}) => {
	const room = `drafts-${Date.now()}`;
	await page.goto(`/?local=1&name=Ada&room=${room}`);
	await expect(page.getByRole('heading', { name: '# General', exact: true })).toBeVisible({
		timeout: 15_000,
	});
	const root = await page.locator('teams-app').evaluate((element) => {
		const client = (element as TeamsApp).client!;
		return client.workspace.chat.post(client.getState().selectedChannelId, {
			text: 'Draft discussion',
		})!.id;
	});
	await page.getByRole('textbox', { name: 'Message', exact: true }).fill('Unsent channel post');
	await page.locator('office-ui-chat-composer input[type=file]').setInputFiles({
		name: 'Budget.xlsx',
		mimeType: 'application/octet-stream',
		buffer: Buffer.from('draft attachment bytes'),
	});
	await page.locator(`[data-message-id="${root}"]`).hover();
	await page
		.locator(`[data-message-id="${root}"]`)
		.getByRole('button', { name: 'Reply', exact: true })
		.click();
	const thread = page.getByRole('complementary', { name: 'Thread', exact: true });
	await thread.getByRole('textbox', { name: 'Message', exact: true }).fill('Unsent thread reply');
	await thread.getByRole('button', { name: 'Close thread', exact: true }).click();
	await expect(page.getByRole('textbox', { name: 'Message', exact: true })).toHaveValue(
		'Unsent channel post',
	);
	await expect(page.getByRole('button', { name: 'Remove Budget.xlsx', exact: true })).toBeVisible();
	expect(
		await page
			.locator('teams-app')
			.evaluate(async (element) => (element as TeamsApp).client!.getState().draft.files[0]!.text()),
	).toBe('draft attachment bytes');
	await page.getByRole('button', { name: 'Drafts', exact: true }).click();
	const drafts = page.getByRole('region', { name: 'Drafts', exact: true });
	await expect(drafts.getByText('Unsent channel post', { exact: true })).toBeVisible();
	await expect(drafts.getByText('Unsent thread reply', { exact: true })).toBeVisible();
	await drafts
		.getByRole('button', { name: 'Resume draft in General: Unsent thread reply' })
		.click();
	await expect(thread.getByRole('textbox', { name: 'Message', exact: true })).toHaveValue(
		'Unsent thread reply',
	);
	await thread.getByRole('button', { name: 'Send message', exact: true }).click();
	await expect(thread.getByText('Unsent thread reply', { exact: true })).toBeVisible();
	await expect
		.poll(() => page.evaluate((room) => localStorage.getItem(`teams:doc:${room}`), room))
		.not.toBeNull();
	await page.reload();
	await expect(page.getByRole('textbox', { name: 'Message', exact: true })).toHaveValue(
		'Unsent channel post',
	);
	await expect(page.getByRole('button', { name: 'Send message', exact: true })).toBeDisabled();
	await expect(
		page.getByRole('button', { name: 'Discard missing Budget.xlsx', exact: true }),
	).toBeVisible();
	await page.getByRole('button', { name: 'Discard missing Budget.xlsx', exact: true }).click();
	await page.getByRole('button', { name: 'Send message', exact: true }).click();
	await expect(page.getByText('Unsent channel post', { exact: true })).toBeVisible();
	await page.getByRole('button', { name: 'Drafts', exact: true }).click();
	await expect(drafts.getByText('No unsent drafts.', { exact: true })).toBeVisible();
});

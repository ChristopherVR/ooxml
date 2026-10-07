import { expect, test } from '@playwright/test';
import type { TeamsApp } from '../../src/ui/src/teams/app/teams-app';

type Attempt = {
	name: string;
	bytes: string;
	signal?: AbortSignal;
	finish: () => void;
	fail: () => void;
};
type TransferWindow = Window & { attempts: Attempt[] };
test('retains failed chat files, cancels ignored uploads and retries in the original thread', async ({
	page,
}) => {
	await page.goto(`/?local=1&name=Ada&room=chat-transfer-${Date.now()}`);
	await expect(page.getByRole('heading', { name: '# General', exact: true })).toBeVisible({
		timeout: 15_000,
	});
	await page.locator('teams-app').evaluate((element) => {
		const app = element as TeamsApp;
		const state = window as TransferWindow;
		state.attempts = [];
		app.uploadFile = (file, context) =>
			new Promise<{ url: string }>((resolve, reject) => {
				const attempt: Attempt = {
					name: file.name,
					bytes: '',
					...(context.signal ? { signal: context.signal } : {}),
					finish: () => resolve({ url: `https://files.test/${file.name}` }),
					fail: () => reject(new Error('Storage offline')),
				};
				state.attempts.push(attempt);
				void (file as Blob).text().then((bytes) => {
					attempt.bytes = bytes;
				});
			});
		app.userId = 'chat-transfer-ada';
	});
	await expect
		.poll(() =>
			page
				.locator('teams-app')
				.evaluate((element) => (element as TeamsApp).client?.getState().user.id),
		)
		.toBe('chat-transfer-ada');
	const root = await page.locator('teams-app').evaluate(async (element) => {
		const client = (element as TeamsApp).client!;
		client.createChannel('Attachments');
		await Promise.resolve();
		return client.workspace.chat.post(client.getState().selectedChannelId, {
			text: 'Original file thread',
		})!.id;
	});
	await page.locator(`[data-message-id="${root}"]`).hover();
	await page
		.locator(`[data-message-id="${root}"]`)
		.getByRole('button', { name: 'Reply', exact: true })
		.click();
	const thread = page.getByRole('complementary', { name: 'Thread', exact: true });
	await thread.getByRole('textbox', { name: 'Message', exact: true }).fill('Attached budget');
	await thread.locator('input[type=file]').setInputFiles([
		{
			name: 'Budget.xlsx',
			mimeType: 'application/octet-stream',
			buffer: Buffer.from('budget bytes'),
		},
		{ name: 'Notes.md', mimeType: 'text/markdown', buffer: Buffer.from('notes bytes') },
	]);
	await thread.getByRole('button', { name: 'Send message', exact: true }).click();
	await expect(page.getByText(/0 of 2 files uploaded/)).toBeVisible();
	await page.evaluate(() => (window as TransferWindow).attempts[0]!.finish());
	await expect(page.getByText(/1 of 2 files uploaded/)).toBeVisible();
	await page.evaluate(() => (window as TransferWindow).attempts[1]!.fail());
	await expect(thread.getByRole('textbox', { name: 'Message', exact: true })).toHaveValue(
		'Attached budget',
	);
	await expect(
		thread.getByRole('button', { name: 'Remove Budget.xlsx', exact: true }),
	).toBeVisible();
	await expect(thread.getByRole('button', { name: 'Remove Notes.md', exact: true })).toBeVisible();
	await expect
		.poll(() =>
			page
				.locator('teams-app')
				.evaluate((element) => (element as TeamsApp).client!.getState().messages.length),
		)
		.toBe(1);
	await thread.getByRole('button', { name: 'Send message', exact: true }).click();
	await expect(
		page.getByRole('button', { name: 'Cancel message send', exact: true }),
	).toBeVisible();
	await thread.getByRole('textbox', { name: 'Message', exact: true }).fill('Newer draft');
	await page.getByRole('button', { name: 'Cancel message send', exact: true }).click();
	await expect(page.getByRole('button', { name: 'Cancel message send', exact: true })).toHaveCount(
		0,
	);
	expect(await page.evaluate(() => (window as TransferWindow).attempts[2]!.signal?.aborted)).toBe(
		true,
	);
	await page.evaluate(() => (window as TransferWindow).attempts[2]!.finish());
	await expect(thread.getByRole('textbox', { name: 'Message', exact: true })).toHaveValue(
		'Newer draft',
	);
	await page.getByRole('button', { name: 'Drafts', exact: true }).click();
	const drafts = page.getByRole('region', { name: 'Drafts', exact: true });
	await expect(drafts.getByText('Newer draft', { exact: true })).toBeVisible();
	await drafts
		.getByRole('button', { name: 'Resume draft in Attachments: Attached budget' })
		.click();
	await thread.getByRole('button', { name: 'Send message', exact: true }).click();
	await expect(page.getByText(/0 of 2 files uploaded/)).toBeVisible();
	await page.evaluate(() => (window as TransferWindow).attempts[3]!.finish());
	await expect(page.getByText(/1 of 2 files uploaded/)).toBeVisible();
	await page.evaluate(() => (window as TransferWindow).attempts[4]!.finish());
	await expect(thread.getByText('Attached budget', { exact: true })).toBeVisible();
	const messages = await page
		.locator('teams-app')
		.evaluate((element) => (element as TeamsApp).client!.getState().messages);
	expect(messages).toHaveLength(2);
	expect(messages[1]).toMatchObject({
		replyTo: root,
		attachments: [{ name: 'Budget.xlsx', kind: 'xlsx' }, { name: 'Notes.md' }],
	});
	const attempts = await page.evaluate(() =>
		(window as TransferWindow).attempts.map(({ name, bytes }) => ({ name, bytes })),
	);
	expect(new Set(attempts.map((attempt) => attempt.name)).size).toBe(5);
	expect(attempts.map((attempt) => attempt.bytes)).toEqual([
		'budget bytes',
		'notes bytes',
		'budget bytes',
		'budget bytes',
		'notes bytes',
	]);
	await page.getByRole('button', { name: 'Drafts', exact: true }).click();
	await expect(drafts.getByText('Newer draft', { exact: true })).toBeVisible();
});

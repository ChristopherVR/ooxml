import { expect, test, type BrowserContext, type Page } from '@playwright/test';
import { join } from 'node:path';
import type { TeamsApp } from '../../src/ui/src/teams/app/teams-app';

// Set TEAMS_SHOTS_DIR to keep screenshots of the Chat app (1366x768 and 375x812).
const shotsDir = process.env.TEAMS_SHOTS_DIR;
const shot = async (page: Page, name: string): Promise<void> => {
	if (shotsDir) await page.screenshot({ path: join(shotsDir, `${name}.png`) });
};

async function open(context: BrowserContext, name: string, room: string): Promise<Page> {
	const page = await context.newPage();
	await page.goto(`/?local=1&name=${name}&room=${room}`);
	await expect(page.getByRole('heading', { name: '# General', exact: true })).toBeVisible({
		timeout: 15_000,
	});
	return page;
}
const client = <T>(page: Page, fn: (c: NonNullable<TeamsApp['client']>) => T): Promise<T> =>
	page.locator('teams-app').evaluate((el, source) => {
		const run = new Function('c', `return (${source})(c)`) as (c: unknown) => unknown;
		return run((el as TeamsApp).client) as never;
	}, fn.toString());
const rail = (page: Page) => page.getByRole('navigation', { name: 'App bar', exact: true });
const railOrder = async (page: Page): Promise<string[]> =>
	(
		await rail(page)
			.getByRole('button')
			.evaluateAll((els) => els.map((e) => e.getAttribute('aria-label') ?? ''))
	).slice(0, 3);
const composer = (page: Page) => page.getByRole('textbox', { name: 'Message', exact: true });
const sidebar = (page: Page) => page.locator('office-ui-chat-sidebar');
const post = async (page: Page, text: string) => {
	await composer(page).fill(text);
	await composer(page).press('Enter');
};
async function startChat(page: Page, people: string[], name = ''): Promise<void> {
	await rail(page).getByRole('button', { name: 'Chat', exact: true }).click();
	await sidebar(page).getByRole('button', { name: 'New chat', exact: true }).click();
	const dialog = page.getByRole('dialog', { name: 'New chat', exact: true });
	await expect(dialog).toBeVisible();
	for (const person of people)
		await dialog.getByRole('checkbox', { name: new RegExp(person, 'u') }).check();
	if (name) await dialog.getByRole('textbox', { name: /Group name/u }).fill(name);
	await dialog.getByRole('button', { name: /^Start (group )?chat$/u }).click();
}

test.use({ viewport: { width: 1366, height: 768 } });

test('1:1 chat: rail position, privacy for a third person, unread badge and replies', async ({
	browser,
}) => {
	const room = `chat11-${Date.now()}`;
	const context = await browser.newContext();
	const ada = await open(context, 'Ada', room);
	const bob = await open(context, 'Bob', room);
	const eve = await open(context, 'Eve', room);
	// Chat is the second app, right after Activity.
	expect(await railOrder(ada)).toEqual(['Activity', 'Chat', 'Teams']);
	await expect.poll(() => client(ada, (c) => c.getState().contacts.length)).toBe(2);

	await startChat(ada, ['Bob']);
	await expect(ada.getByRole('heading', { name: 'Bob', exact: true })).toBeVisible();
	await expect(composer(ada)).toHaveAttribute('placeholder', 'Message Bob');
	await post(ada, 'Only for you, Bob');
	await shot(ada, 'chat-1366x768-conversation');

	// Bob gets the chat with an unread badge on the rail and the row, then reads and replies.
	await expect(sidebar(bob)).toHaveCount(0);
	await expect(rail(bob).getByRole('button', { name: 'Chat', exact: true })).toContainText('1');
	await rail(bob).getByRole('button', { name: 'Chat', exact: true }).click();
	const row = sidebar(bob).getByRole('button', { name: /Ada/u });
	await expect(row).toContainText('Only for you, Bob');
	await expect(row.locator('.count')).toHaveText('1');
	await shot(bob, 'chat-1366x768-list-unread');
	await row.click();
	await expect(
		bob.locator('office-ui-chat-list .msg', { hasText: 'Only for you, Bob' }),
	).toBeVisible();
	await expect(row.locator('.count')).toHaveCount(0);
	await post(bob, 'Got it, thanks');
	await expect(
		ada.locator('office-ui-chat-list .msg', { hasText: 'Got it, thanks' }),
	).toBeVisible();

	// Reactions, mentions and pins work like in a channel.
	await ada.locator('office-ui-chat-list .msg', { hasText: 'Got it, thanks' }).hover();
	await ada.getByRole('button', { name: 'React 👍', exact: true }).first().click();
	await expect(bob.locator('office-ui-chat-list .pill').first()).toContainText('👍');
	await bob.locator('office-ui-chat-list .msg', { hasText: 'Only for you, Bob' }).hover();
	await bob.getByRole('button', { name: 'Pin', exact: true }).first().click();
	await expect(bob.locator('teams-pinned-bar')).toContainText('Only for you, Bob');

	// Eve is in the workspace but not in the chat: no row, no messages, no way in.
	await eve.waitForTimeout(300);
	const chatId = await client(ada, (c) => c.getState().chats[0]!.id);
	expect(await client(eve, (c) => c.getState().chats.length)).toBe(0);
	await rail(eve).getByRole('button', { name: 'Chat', exact: true }).click();
	await expect(sidebar(eve)).toContainText('No chats yet');
	await expect(eve.getByText('Only for you, Bob')).toHaveCount(0);
	const denied = await eve.locator('teams-app').evaluate((el, id) => {
		const c = (el as TeamsApp).client!;
		const before = c.getState().selectedChannelId;
		c.select(id);
		return {
			messages: c.workspace.chat.messages(id).length,
			posted: c.workspace.chat.post(id, { text: 'let me in' }),
			selectedChanged: c.getState().selectedChannelId === id && before !== id,
		};
	}, chatId);
	expect(denied).toEqual({ messages: 0, posted: null, selectedChanged: false });
	await expect(ada.getByText('let me in')).toHaveCount(0);
	await context.close();
});

test('group chat: name, people dialog, system notices, leave', async ({ browser }) => {
	const room = `chatgrp-${Date.now()}`;
	const context = await browser.newContext();
	const ada = await open(context, 'Ada', room);
	const bob = await open(context, 'Bob', room);
	const cy = await open(context, 'Cy', room);
	await expect.poll(() => client(ada, (c) => c.getState().contacts.length)).toBe(2);
	await startChat(ada, ['Bob', 'Cy'], 'Launch team');
	await expect(ada.getByRole('heading', { name: 'Launch team', exact: true })).toBeVisible();
	await post(ada, 'Welcome to the launch chat');
	for (const page of [bob, cy]) {
		await rail(page).getByRole('button', { name: 'Chat', exact: true }).click();
		await expect(sidebar(page).getByRole('button', { name: /Launch team/u })).toBeVisible();
	}
	await sidebar(bob)
		.getByRole('button', { name: /Launch team/u })
		.click();
	await expect(
		bob.locator('office-ui-chat-list .msg', { hasText: 'Welcome to the launch chat' }),
	).toBeVisible();

	// Rename and add a person from the People dialog; the change shows as a notice line.
	await ada.getByRole('button', { name: 'People', exact: true }).click();
	const people = ada.getByRole('dialog', { name: 'People', exact: true });
	await expect(people).toBeVisible();
	await people.getByRole('textbox', { name: /Group name/u }).fill('Launch crew');
	await people.getByRole('button', { name: 'Rename', exact: true }).click();
	await expect(
		ada.locator('office-ui-chat-list .system', { hasText: 'named the chat' }),
	).toBeVisible();
	await shot(ada, 'chat-1366x768-people');
	await people.getByRole('button', { name: 'Remove Cy', exact: true }).click();
	await expect(ada.locator('office-ui-chat-list .system', { hasText: 'removed Cy' })).toBeVisible();
	await people.getByRole('button', { name: 'Done', exact: true }).click();

	// Cy no longer has the chat or its messages; Bob leaves on his own.
	await expect.poll(() => client(cy, (c) => c.getState().chats.length)).toBe(0);
	bob.once('dialog', (dialog) => void dialog.accept());
	await bob.getByRole('button', { name: 'People', exact: true }).click();
	await bob
		.getByRole('dialog', { name: 'People', exact: true })
		.getByRole('button', { name: 'Leave chat' })
		.click();
	await expect.poll(() => client(bob, (c) => c.getState().chats.length)).toBe(0);
	await expect(
		ada.locator('office-ui-chat-list .system', { hasText: 'Bob left the chat' }),
	).toBeVisible();
	await context.close();
});

test('phone: chat list and conversation are full screen views with the bottom bar', async ({
	browser,
}) => {
	const room = `chatphone-${Date.now()}`;
	const context = await browser.newContext({ viewport: { width: 375, height: 812 } });
	const ada = await open(context, 'Ada', room);
	const bob = await open(context, 'Bob', room);
	await expect.poll(() => client(ada, (c) => c.getState().contacts.length)).toBe(1);
	const bar = rail(ada);
	const box = (await bar.boundingBox())!;
	expect(box.y + box.height).toBeGreaterThan(800);
	// The phone bar shows Activity, Chat, Teams, Calendar and More; the rest sit in the More sheet.
	const visible = bar.locator('ul').first().locator('li:not([data-secondary="true"]) button');
	await expect(visible).toHaveCount(5);
	await expect(visible).toHaveText([/Activity/u, /Chat/u, /Teams/u, /Calendar/u, /More/u]);
	for (const item of await visible.all()) {
		const b = (await item.boundingBox())!;
		expect(b.x).toBeGreaterThanOrEqual(0);
		expect(b.x + b.width).toBeLessThanOrEqual(375);
	}
	await expect(bar.getByRole('button', { name: 'Files', exact: true })).toBeHidden();
	const more = bar.getByRole('button', { name: 'More', exact: true });
	await more.click();
	const sheet = ada.getByRole('dialog', { name: 'More apps' });
	await expect(sheet).toBeVisible();
	await expect(sheet.getByRole('button')).toHaveText([
		/Calls/u,
		/Files/u,
		/Followed threads/u,
		/Drafts/u,
		/Close/u,
	]);
	await shot(ada, 'phone-375x812-more-sheet');
	await ada.keyboard.press('Escape');
	await expect(sheet).toBeHidden();
	await expect(more).toBeFocused();
	await more.click();
	await sheet.getByRole('button', { name: 'Files', exact: true }).click();
	await expect(sheet).toBeHidden();
	await expect(ada.getByRole('heading', { name: 'Files', exact: true })).toBeVisible();
	await expect(more).toHaveAttribute('aria-current', 'page');
	await bar.getByRole('button', { name: 'Chat', exact: true }).click();
	await expect(sidebar(ada)).toBeVisible();
	await sidebar(ada).getByRole('button', { name: 'New chat', exact: true }).click();
	const dialog = ada.getByRole('dialog', { name: 'New chat', exact: true });
	const dbox = (await dialog.boundingBox())!;
	expect(dbox.width).toBeGreaterThanOrEqual(370);
	await shot(ada, 'chat-375x812-new-chat');
	await dialog.getByRole('checkbox', { name: /Bob/u }).check();
	await dialog.getByRole('button', { name: 'Start chat', exact: true }).click();
	await expect(ada.getByRole('heading', { name: 'Bob', exact: true })).toBeVisible();
	await expect(sidebar(ada)).toBeHidden();
	await post(ada, 'Hello from the phone');
	await shot(ada, 'chat-375x812-conversation');
	await ada.getByRole('button', { name: 'Back to chats', exact: true }).click();
	await expect(sidebar(ada)).toBeVisible();
	await expect(sidebar(ada)).toContainText('Hello from the phone');
	await shot(ada, 'chat-375x812-list');
	await bob.close();
	await context.close();
});

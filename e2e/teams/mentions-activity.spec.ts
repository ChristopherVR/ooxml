import { expect, test, type BrowserContext, type Page } from '@playwright/test';
import type { TeamsApp } from '../../src/ui/src/teams/app/teams-app';

declare global {
	interface Window {
		__hidden?: boolean;
		__notes?: { title: string; body: string }[];
	}
}

async function open(context: BrowserContext, name: string, room: string): Promise<Page> {
	const page = await context.newPage();
	await page.goto(`/?local=1&name=${name}&room=${room}`);
	await expect(page.getByRole('heading', { name: '# General', exact: true })).toBeVisible({
		timeout: 15_000,
	});
	return page;
}
const composer = (page: Page) => page.getByRole('textbox', { name: 'Message', exact: true });
const rail = (page: Page) => page.getByRole('navigation', { name: 'App bar', exact: true });
const activityButton = (page: Page) => rail(page).getByRole('button', { name: 'Activity' });
const post = async (page: Page, text: string) => {
	await composer(page).fill(text);
	await composer(page).press('Enter');
};
const row = (page: Page, text: string) =>
	page.locator('office-ui-chat-list .msg', { hasText: text }).first();

test('mentions: picker, highlighting, emphasis for you, and edits keep them', async ({
	browser,
}) => {
	const room = `mentions-${Date.now()}`;
	const context = await browser.newContext();
	const ada = await open(context, 'Ada', room);
	const bob = await open(context, 'Bob', room);
	await post(ada, 'Morning all');
	await expect(row(bob, 'Morning all')).toBeVisible();

	await composer(bob).pressSequentially('hey @a');
	const picker = bob.getByRole('listbox', { name: 'People to mention' });
	await expect(picker).toBeVisible();
	await expect(picker.getByRole('option', { name: /Ada/u })).toBeVisible();
	await bob.screenshot({ path: test.info().outputPath('mention-picker.png') });
	await composer(bob).press('Enter');
	await expect(composer(bob)).toHaveValue('hey Ada ');
	await expect(picker).toBeHidden();
	await composer(bob).pressSequentially('please look');
	await composer(bob).press('Enter');

	const forAda = ada.locator('office-ui-chat-list .mention[data-self="true"]');
	await expect(forAda).toHaveText('Ada');
	await expect(ada.locator('office-ui-chat-list .msg[data-mentions-you="true"]')).toHaveCount(1);
	await expect(bob.locator('office-ui-chat-list .mention[data-self="false"]')).toHaveText('Ada');
	await ada.screenshot({ path: test.info().outputPath('mention-you.png') });

	// @channel is offered, and editing the message keeps the mention.
	await composer(bob).pressSequentially('@ch');
	await expect(bob.getByRole('option', { name: /channel/u })).toBeVisible();
	await composer(bob).press('Escape');
	await expect(picker).toBeHidden();
	await composer(bob).fill('');
	const mine = row(bob, 'please look');
	await mine.hover();
	await mine.getByRole('button', { name: 'Edit', exact: true }).click();
	const edit = composer(bob);
	await expect(edit).toHaveValue('hey Ada please look');
	await edit.fill('hey Ada please look now');
	await edit.press('Enter');
	await expect(ada.locator('office-ui-chat-list .mention[data-self="true"]')).toHaveText('Ada');
	await expect(row(ada, 'please look now')).toContainText('(edited)');
	await context.close();
});

test('activity: feed, filters, read state and opening an item', async ({ browser }) => {
	const room = `activity-${Date.now()}`;
	const context = await browser.newContext();
	const ada = await open(context, 'Ada', room);
	const bob = await open(context, 'Bob', room);
	await post(ada, 'Ship the report');
	const bobsView = row(bob, 'Ship the report');
	await expect(bobsView).toBeVisible();
	await bobsView.hover();
	await bobsView.getByRole('button', { name: 'React 👍', exact: true }).click();
	await bobsView.getByRole('button', { name: 'Pin', exact: true }).click();
	await bobsView.getByRole('button', { name: 'Reply', exact: true }).click();
	await composer(bob).fill('On it');
	await composer(bob).press('Enter');
	await post(ada, 'Thanks');
	await bob.getByRole('button', { name: 'Close thread' }).click();
	await post(bob, 'FYI @Ada the report is due Friday');

	await expect(activityButton(ada)).toContainText('4');
	await activityButton(ada).click();
	const feed = ada.getByRole('list', { name: 'Activity feed' });
	await expect(feed.getByRole('listitem')).toHaveCount(4);
	await expect(feed.getByText('Bob mentioned you in # General')).toBeVisible();
	await expect(feed.getByText('Bob replied in # General')).toBeVisible();
	await expect(feed.getByText('Bob reacted 👍 to your message')).toBeVisible();
	await expect(feed.getByText('Bob pinned a message in # General')).toBeVisible();
	await ada.setViewportSize({ width: 1366, height: 768 });
	await ada.screenshot({ path: test.info().outputPath('activity-1366.png') });

	await ada
		.getByRole('group', { name: 'Filter activity' })
		.getByRole('button', { name: 'Mentions' })
		.click();
	await expect(feed.getByRole('listitem')).toHaveCount(1);
	await ada.getByRole('button', { name: /^All$/u }).click();

	await feed.getByRole('button', { name: /replied in # General/u }).click();
	await expect(ada.getByRole('complementary', { name: 'Thread' })).toContainText('On it');
	await expect(activityButton(ada)).toContainText('3');

	await ada.getByRole('button', { name: 'Mark all as read' }).click();
	await expect(rail(ada).locator('.badge')).toHaveCount(0);
	await expect(ada.getByRole('button', { name: 'Mark all as read' })).toBeDisabled();
	await ada.getByRole('button', { name: /^Unread/u }).click();
	await expect(ada.getByText('You are all caught up.')).toBeVisible();

	// The mention opens the post itself, scrolled into view.
	await ada.getByRole('button', { name: /^All$/u }).click();
	await feed.getByRole('button', { name: /mentioned you in # General/u }).click();
	await expect(row(ada, 'FYI Ada')).toBeInViewport();

	// Phone layout: the feed fills the screen and an item opens the conversation.
	await ada.setViewportSize({ width: 375, height: 812 });
	await activityButton(ada).click();
	await expect(feed).toBeVisible();
	await ada.screenshot({ path: test.info().outputPath('activity-375.png') });
	await feed.getByRole('button', { name: /mentioned you in # General/u }).click();
	await expect(row(ada, 'FYI Ada')).toBeInViewport();
	await ada.screenshot({ path: test.info().outputPath('activity-open-375.png') });
	// A toast on a phone is a full-width banner below the header.
	await post(bob, 'ping @Ada on the phone');
	const phoneToast = ada.locator('.activity-toast').first();
	await expect(phoneToast).toContainText('ping Ada on the phone');
	const banner = (await phoneToast.boundingBox())!;
	expect(banner.width).toBeGreaterThan(340);
	expect(banner.y).toBeGreaterThan(50);
	await ada.screenshot({ path: test.info().outputPath('activity-toast-375.png') });
	await context.close();
});

test('notifications: Do not disturb suppresses them and settings choose the level', async ({
	browser,
}) => {
	const room = `notify-${Date.now()}`;
	const context = await browser.newContext();
	await context.addInitScript(() => {
		window.__notes = [];
		class FakeNotification {
			static permission = 'granted';
			static requestPermission = () => Promise.resolve('granted');
			onclick: (() => void) | null = null;
			constructor(title: string, options?: { body?: string }) {
				window.__notes!.push({ title, body: options?.body ?? '' });
			}
			close(): void {}
		}
		Object.defineProperty(window, 'Notification', { value: FakeNotification, configurable: true });
		Object.defineProperty(document, 'visibilityState', {
			configurable: true,
			get: () => (window.__hidden ? 'hidden' : 'visible'),
		});
	});
	const ada = await open(context, 'Ada', room);
	const bob = await open(context, 'Bob', room);
	await post(bob, 'Hello');
	await expect(row(ada, 'Hello')).toBeVisible();
	const toast = ada.locator('.activity-toast');

	// Do not disturb: the mention lands in Activity, but nothing is announced.
	await ada.getByRole('button', { name: 'Profile for Ada', exact: true }).click();
	const profile = ada.getByRole('dialog', { name: 'Your profile', exact: true });
	await profile.getByRole('combobox', { name: 'Availability', exact: true }).selectOption('dnd');
	await expect
		.poll(() =>
			ada.locator('teams-app').evaluate((el) => (el as TeamsApp).client!.getState().availability),
		)
		.toBe('dnd');
	await ada.keyboard.press('Escape');
	await post(bob, 'psst @Ada are you there');
	await expect(activityButton(ada)).toContainText('1');
	await ada.waitForTimeout(1500);
	await expect(toast).toHaveCount(0);
	expect(await ada.evaluate(() => window.__notes!.length)).toBe(0);

	// Available again: a visible tab shows a toast.
	await ada.getByRole('button', { name: 'Profile for Ada', exact: true }).click();
	await profile
		.getByRole('combobox', { name: 'Availability', exact: true })
		.selectOption('available');
	await ada.keyboard.press('Escape');
	await post(bob, 'again @Ada please');
	await expect(toast).toContainText('Bob mentioned you in # General');
	await expect(toast).toContainText('again Ada please');
	await ada.screenshot({ path: test.info().outputPath('notification-toast.png') });
	// Several toasts stack, and none covers the compose box.
	await post(bob, 'and @Ada once more');
	await expect(toast).toHaveCount(2);
	const toastBox = (await toast.first().boundingBox())!;
	const lastBox = (await toast.last().boundingBox())!;
	const composerBox = (await composer(ada).boundingBox())!;
	expect(toastBox.y + toastBox.height).toBeLessThan(lastBox.y + 1);
	expect(lastBox.y + lastBox.height).toBeLessThan(composerBox.y);
	await ada.screenshot({ path: test.info().outputPath('notification-toasts-stacked.png') });
	await toast.first().getByRole('button', { name: 'Dismiss' }).click();
	await toast.first().getByRole('button', { name: 'View' }).click();
	await expect(toast).toHaveCount(0);

	// A hidden tab gets a desktop notification instead of a toast.
	await ada.evaluate(() => {
		window.__hidden = true;
	});
	await post(bob, 'third @Ada ping');
	await expect.poll(() => ada.evaluate(() => window.__notes!.length)).toBe(1);
	expect(await ada.evaluate(() => window.__notes![0]?.title)).toBe(
		'Bob mentioned you in # General',
	);
	await expect(toast).toHaveCount(0);
	await ada.evaluate(() => {
		window.__hidden = false;
	});

	// Settings: mentions only silences reactions; off silences everything.
	await ada.getByRole('button', { name: 'Settings and more', exact: true }).click();
	await ada.getByRole('menuitem', { name: 'Settings', exact: true }).click();
	await ada.getByRole('tab', { name: 'Notifications and activity', exact: true }).click();
	await expect(ada.getByRole('tabpanel', { name: 'Notifications and activity' })).not.toContainText(
		'not available yet',
	);
	await ada.screenshot({ path: test.info().outputPath('notification-settings.png') });
	await ada.getByRole('radio', { name: /Mentions only/u }).check();
	await expect
		.poll(() =>
			ada
				.locator('teams-app')
				.evaluate((el) => (el as TeamsApp).client!.getState().notificationSettings.level),
		)
		.toBe('mentions');
	await ada.getByRole('button', { name: 'Close settings', exact: true }).click();
	await post(ada, 'My own post');
	const theirs = row(bob, 'My own post');
	await theirs.hover();
	await theirs.getByRole('button', { name: 'React 👍', exact: true }).click();
	await expect(activityButton(ada)).toContainText('4');
	await ada.waitForTimeout(1000);
	await expect(toast).toHaveCount(0);
	await context.close();
});

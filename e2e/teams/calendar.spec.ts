import { expect, test, type BrowserContext, type Page } from '@playwright/test';
import { readFile } from 'node:fs/promises';
import { join } from 'node:path';
import type { TeamsApp } from '../../src/ui/src/teams/app/teams-app';

// Set TEAMS_SHOTS_DIR to keep screenshots of the Calendar app (1366x768 and 375x812).
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
const rail = (page: Page) => page.getByRole('navigation', { name: 'App bar', exact: true });
const calendar = (page: Page) => page.locator('office-ui-calendar-view');
const agenda = (page: Page) => page.locator('office-ui-calendar-agenda');
const count = (page: Page, expr: 'contacts' | 'meetings' | 'chats'): Promise<number> =>
	page.locator('teams-app').evaluate((el, key) => {
		const state = (el as TeamsApp).client!.getState();
		return state[key as 'contacts'].length;
	}, expr);
/** The `datetime-local` value `minutes` from now, rounded down to the minute, in this browser's zone. */
const inMinutes = (minutes: number): string => {
	const d = new Date(Date.now() + minutes * 60_000);
	const p = (n: number): string => String(n).padStart(2, '0');
	return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}T${p(d.getHours())}:${p(d.getMinutes())}`;
};

async function scheduleMeeting(
	page: Page,
	options: {
		title: string;
		startIn: number;
		lengthMinutes?: number;
		attendees?: string[];
		repeat?: string;
		channel?: string;
	},
): Promise<void> {
	await rail(page).getByRole('button', { name: 'Calendar', exact: true }).click();
	await page.getByRole('button', { name: 'New meeting', exact: true }).first().click();
	const dialog = page.getByRole('dialog', { name: 'New meeting', exact: true });
	await expect(dialog).toBeVisible();
	await dialog.getByRole('textbox', { name: 'Title', exact: true }).fill(options.title);
	await dialog.getByLabel('Start', { exact: true }).fill(inMinutes(options.startIn));
	await dialog
		.getByLabel('End', { exact: true })
		.fill(inMinutes(options.startIn + (options.lengthMinutes ?? 30)));
	if (options.repeat)
		await dialog.getByRole('combobox', { name: /^Repeat/u }).selectOption(options.repeat);
	if (options.channel)
		await dialog
			.getByRole('combobox', { name: /^Channel/u })
			.selectOption({ label: options.channel });
	for (const person of options.attendees ?? [])
		await dialog.getByRole('checkbox', { name: new RegExp(person, 'u') }).check();
	await dialog.getByRole('button', { name: 'Send', exact: true }).click();
	await expect(dialog).toBeHidden();
}

test.use({ viewport: { width: 1366, height: 768 } });

test('calendar: week view, scheduling, details, join and .ics', async ({ browser }) => {
	const room = `cal-${Date.now()}`;
	const context = await browser.newContext({ acceptDownloads: true });
	const ada = await open(context, 'Ada', room);
	const bob = await open(context, 'Bob', room);
	const eve = await open(context, 'Eve', room);
	await expect.poll(() => count(ada, 'contacts')).toBe(2);

	// Calendar follows Teams in the rail and opens a work week with Today highlighted.
	expect(
		await rail(ada)
			.getByRole('button')
			.evaluateAll((els) => els.map((e) => e.getAttribute('aria-label')).slice(2, 5)),
	).toEqual(['Teams', 'Calendar', 'Calls']);
	await rail(ada).getByRole('button', { name: 'Calendar', exact: true }).click();
	await expect(calendar(ada)).toBeVisible();
	await expect(calendar(ada).locator('.column')).toHaveCount(5);
	await expect(
		calendar(ada)
			.getByRole('group', { name: 'Calendar view' })
			.getByRole('button', { pressed: true }),
	).toHaveText('Work week');
	await calendar(ada).getByRole('button', { name: 'Week', exact: true }).click();
	await expect(calendar(ada).locator('.column')).toHaveCount(7);
	await calendar(ada).getByRole('button', { name: 'Day', exact: true }).click();
	await expect(calendar(ada).locator('.column')).toHaveCount(1);
	await calendar(ada).getByRole('button', { name: 'Next day', exact: true }).click();
	await calendar(ada).getByRole('button', { name: 'Today', exact: true }).click();
	await calendar(ada).getByRole('button', { name: 'Work week', exact: true }).click();

	// Schedule a repeating meeting with Bob; Eve is not invited.
	await scheduleMeeting(ada, {
		title: 'Design sync',
		startIn: 3,
		lengthMinutes: 45,
		attendees: ['Bob'],
		repeat: 'daily',
	});
	await expect.poll(() => count(bob, 'meetings')).toBe(1);
	expect(await count(eve, 'meetings')).toBe(0);
	await expect(agenda(ada)).toContainText('Design sync');
	await shot(ada, 'calendar-1366x768-week');

	// Bob sees it in his own calendar, gets the meeting chat with its card and can join from it.
	await rail(bob).getByRole('button', { name: 'Calendar', exact: true }).click();
	await expect(agenda(bob)).toContainText('Design sync');
	await agenda(bob)
		.getByRole('button', { name: /^Design sync/u })
		.first()
		.click();
	const details = bob.getByRole('dialog', { name: 'Design sync', exact: true });
	await expect(details).toBeVisible();
	await expect(details).toContainText('Repeats every day');
	await expect(details).toContainText('Organized by Ada');
	await expect(details.getByRole('button', { name: 'Edit', exact: true })).toHaveCount(0);
	await shot(bob, 'calendar-1366x768-details');
	const [download] = await Promise.all([
		bob.waitForEvent('download'),
		details.getByRole('button', { name: 'Export .ics', exact: true }).click(),
	]);
	expect(download.suggestedFilename()).toBe('Design-sync.ics');
	const ics = await readFile((await download.path())!, 'utf8');
	expect(ics).toContain('BEGIN:VEVENT');
	expect(ics).toContain('SUMMARY:Design sync');
	expect(ics).toContain('RRULE:FREQ=DAILY');
	await details.getByRole('button', { name: 'Meeting chat', exact: true }).click();
	const card = bob.locator('office-ui-chat-list .meeting-card');
	await expect(card).toContainText('Design sync');
	await card.getByRole('button', { name: /^Join/u }).click();
	await expect(bob.locator('office-ui-prejoin')).toBeVisible();
	await expect(bob.locator('office-ui-prejoin')).toHaveAttribute('heading', 'Design sync');
	await bob.getByRole('button', { name: 'Cancel', exact: true }).click();

	// The organizer edits and then cancels; the meeting leaves everyone's calendar.
	await calendar(ada)
		.getByRole('button', { name: /^Design sync/u })
		.first()
		.click();
	const own = ada.getByRole('dialog', { name: 'Design sync', exact: true });
	await own.getByRole('button', { name: 'Edit', exact: true }).click();
	const form = ada.getByRole('dialog', { name: 'Edit meeting', exact: true });
	await form.getByRole('textbox', { name: 'Title', exact: true }).fill('Design sync v2');
	await form.getByRole('button', { name: 'Save', exact: true }).click();
	await expect(agenda(bob)).toContainText('Design sync v2');
	ada.once('dialog', (dialog) => void dialog.accept());
	await calendar(ada)
		.getByRole('button', { name: /^Design sync v2/u })
		.first()
		.click();
	await ada
		.getByRole('dialog', { name: 'Design sync v2', exact: true })
		.getByRole('button', { name: 'Cancel meeting' })
		.click();
	await expect.poll(() => count(bob, 'meetings')).toBe(0);
	await context.close();
});

test('calendar: a channel meeting posts a card in the channel and a call can start from it', async ({
	browser,
}) => {
	const room = `calch-${Date.now()}`;
	const context = await browser.newContext();
	const ada = await open(context, 'Ada', room);
	const bob = await open(context, 'Bob', room);
	await scheduleMeeting(ada, { title: 'All hands', startIn: 60, channel: '# General' });
	const card = bob.locator('office-ui-chat-list .meeting-card');
	await expect(card).toContainText('All hands');
	await expect.poll(() => count(bob, 'meetings')).toBe(1);
	await card.getByRole('button', { name: 'Details', exact: true }).click();
	await expect(bob.getByRole('dialog', { name: 'All hands', exact: true })).toContainText(
		'Channel meeting',
	);
	await context.close();
});

test('calendar: a reminder shows five minutes ahead, but not in Do not disturb', async ({
	browser,
}) => {
	test.setTimeout(90_000);
	const room = `calrem-${Date.now()}`;
	const context = await browser.newContext();
	const ada = await open(context, 'Ada', room);
	await scheduleMeeting(ada, { title: 'Standup', startIn: 40 });
	await expect(ada.locator('.activity-toast')).toHaveCount(0);
	// A second meeting inside the window: the 15 second check announces it with a toast.
	await scheduleMeeting(ada, { title: 'Starting soon', startIn: 4 });
	const toast = ada.locator('.activity-toast', { hasText: 'Starting soon' });
	await expect(toast).toBeVisible({ timeout: 25_000 });
	await expect(toast).toContainText('Starts in');
	await shot(ada, 'calendar-1366x768-reminder');
	await toast.getByRole('button', { name: 'View', exact: true }).click();
	await expect(ada.getByRole('dialog', { name: 'Starting soon', exact: true })).toBeVisible();

	// Do not disturb holds the next reminder back.
	await ada.getByRole('button', { name: 'Close meeting details', exact: true }).click();
	await ada.evaluate(() =>
		(document.querySelector('teams-app') as unknown as TeamsApp).client!.setAvailability('dnd'),
	);
	await scheduleMeeting(ada, { title: 'Quiet one', startIn: 3 });
	await ada.waitForTimeout(18_000);
	await expect(ada.locator('.activity-toast', { hasText: 'Quiet one' })).toHaveCount(0);
	await context.close();
});

test('phone: calendar list, details and week view are full screen views', async ({ browser }) => {
	const room = `calphone-${Date.now()}`;
	const context = await browser.newContext({ viewport: { width: 375, height: 812 } });
	const ada = await open(context, 'Ada', room);
	await rail(ada).getByRole('button', { name: 'Calendar', exact: true }).click();
	await expect(agenda(ada)).toBeVisible();
	await expect(calendar(ada)).toBeHidden();
	await agenda(ada).getByRole('button', { name: 'New meeting', exact: true }).click();
	const form = ada.getByRole('dialog', { name: 'New meeting', exact: true });
	const box = (await form.boundingBox())!;
	expect(box.width).toBeGreaterThanOrEqual(370);
	expect(box.height).toBeGreaterThanOrEqual(800);
	await form.getByRole('textbox', { name: 'Title', exact: true }).fill('Phone meeting');
	await form.getByLabel('Start', { exact: true }).fill(inMinutes(90));
	await form.getByLabel('End', { exact: true }).fill(inMinutes(120));
	await shot(ada, 'calendar-375x812-new-meeting');
	await form.getByRole('button', { name: 'Send', exact: true }).click();
	await expect(agenda(ada)).toContainText('Phone meeting');
	await shot(ada, 'calendar-375x812-agenda');
	await agenda(ada)
		.getByRole('button', { name: /^Phone meeting/u })
		.click();
	const details = ada.getByRole('dialog', { name: 'Phone meeting', exact: true });
	await expect(details).toBeVisible();
	await shot(ada, 'calendar-375x812-details');
	await details.getByRole('button', { name: 'Close meeting details', exact: true }).click();
	await ada.getByRole('button', { name: 'Open calendar', exact: true }).click();
	await expect(calendar(ada)).toBeVisible();
	await expect(agenda(ada)).toBeHidden();
	await expect(calendar(ada).locator('.column')).toHaveCount(1);
	await shot(ada, 'calendar-375x812-day');
	expect(await ada.evaluate(() => document.documentElement.scrollWidth)).toBeLessThanOrEqual(375);
	await ada.getByRole('button', { name: 'Upcoming meetings', exact: true }).click();
	await expect(agenda(ada)).toBeVisible();
	await context.close();
});

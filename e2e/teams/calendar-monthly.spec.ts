import { expect, test, type Page } from '@playwright/test';
import { readFile } from 'node:fs/promises';

const inMinutes = (minutes: number): string => {
	const d = new Date(Date.now() + minutes * 60_000);
	const p = (n: number): string => String(n).padStart(2, '0');
	return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}T${p(d.getHours())}:${p(d.getMinutes())}`;
};

async function schedule(page: Page, title: string, repeat: { index: number } | { value: string }) {
	await page
		.getByRole('navigation', { name: 'App bar', exact: true })
		.getByRole('button', { name: 'Calendar', exact: true })
		.click();
	await page.getByRole('button', { name: 'New meeting', exact: true }).first().click();
	const dialog = page.getByRole('dialog', { name: 'New meeting', exact: true });
	await dialog.getByRole('textbox', { name: 'Title', exact: true }).fill(title);
	await dialog.getByLabel('Start', { exact: true }).fill(inMinutes(3));
	await dialog.getByLabel('End', { exact: true }).fill(inMinutes(33));
	await dialog.getByRole('combobox', { name: /^Repeat/u }).selectOption(repeat);
	await dialog.getByRole('button', { name: 'Send', exact: true }).click();
	await expect(dialog).toBeHidden();
}

async function exportIcs(page: Page, title: string, sentence: RegExp): Promise<string> {
	await page
		.locator('office-ui-calendar-agenda')
		.getByRole('button', { name: new RegExp(`^${title}`, 'u') })
		.first()
		.click();
	const details = page.getByRole('dialog', { name: title, exact: true });
	await expect(details).toContainText(sentence);
	const [download] = await Promise.all([
		page.waitForEvent('download'),
		details.getByRole('button', { name: 'Export .ics', exact: true }).click(),
	]);
	return readFile((await download.path())!, 'utf8');
}

test.use({ viewport: { width: 1366, height: 768 }, acceptDownloads: true });

test('calendar: monthly weekday and yearly repeats', async ({ page }) => {
	await page.goto(`/?local=1&name=Ada&room=monthly-${Date.now()}`);
	await expect(page.getByRole('heading', { name: '# General', exact: true })).toBeVisible({
		timeout: 15_000,
	});
	// The picker lists the date, the weekday rule and the year for the chosen start.
	await page
		.getByRole('navigation', { name: 'App bar', exact: true })
		.getByRole('button', { name: 'Calendar', exact: true })
		.click();
	await page.getByRole('button', { name: 'New meeting', exact: true }).first().click();
	const options = page
		.getByRole('dialog', { name: 'New meeting', exact: true })
		.getByRole('combobox', { name: /^Repeat/u })
		.locator('option');
	await expect(options.filter({ hasText: /^Every month on day \d+$/u })).toHaveCount(1);
	await expect(
		options.filter({ hasText: /^Every month on the \w+ \w+day$/u }).first(),
	).toBeAttached();
	await expect(options.filter({ hasText: /^Every year on \w+ \d+$/u })).toHaveCount(1);
	await page.keyboard.press('Escape');

	await schedule(page, 'Board review', { index: 4 });
	const weekday = await exportIcs(page, 'Board review', /Repeats every month on the /u);
	expect(weekday).toMatch(/RRULE:FREQ=MONTHLY;BYDAY=-?\d[A-Z]{2}/u);
	await page.keyboard.press('Escape');

	await schedule(page, 'Anniversary', { value: 'yearly' });
	const yearly = await exportIcs(page, 'Anniversary', /Repeats every year on /u);
	expect(yearly).toMatch(/RRULE:FREQ=YEARLY;BYMONTH=\d+;BYMONTHDAY=\d+/u);
});

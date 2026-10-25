import { expect, test } from '@playwright/test';

// `?locale=` reaches <teams-app> through the selected framework binding (the demos pass it as the
// `locale` prop), so this spec runs in every framework of `test:browser:bindings`.

const room = (): string => `locale-${Date.now()}`;

test('the locale prop translates the app shell, the rail and the composer', async ({ page }) => {
	await page.goto(`/?local=1&name=Ada&room=${room()}&locale=de`);
	await expect(page.getByRole('heading', { name: '# General', exact: true })).toBeVisible({
		timeout: 15_000,
	});
	await expect(page.getByRole('navigation', { name: 'App-Leiste', exact: true })).toBeVisible();
	await expect(page.getByRole('button', { name: 'Aktivität', exact: true })).toBeVisible();
	await expect(page.getByRole('button', { name: 'Kalender', exact: true })).toBeVisible();
	await expect(page.getByRole('textbox', { name: 'Nachricht', exact: true })).toBeVisible();
	await expect(
		page.getByRole('button', { name: 'Einstellungen und mehr', exact: true }),
	).toBeVisible();
	await expect(page.getByRole('button', { name: 'Settings and more' })).toHaveCount(0);
});

test('Simplified Chinese, French and Spanish ship too, and an unknown tag falls back to English', async ({
	page,
}) => {
	for (const [locale, name] of [
		['zh-CN', '应用栏'],
		['fr', "Barre d'applications"],
		['es', 'Barra de aplicaciones'],
		['en-GB', 'App bar'],
		['xx', 'App bar'],
	] as const) {
		await page.goto(`/?local=1&name=Ada&room=${room()}&locale=${locale}`);
		await expect(page.getByRole('navigation', { name, exact: true })).toBeVisible({
			timeout: 15_000,
		});
	}
});

test('a translated dialog and a recurrence label come from the dictionary', async ({ page }) => {
	await page.goto(`/?local=1&name=Ada&room=${room()}&locale=de`);
	await page.getByRole('button', { name: 'Kalender', exact: true }).click();
	await page.getByRole('button', { name: 'Neue Besprechung', exact: true }).first().click();
	const dialog = page.getByRole('dialog', { name: 'Neue Besprechung' });
	await expect(dialog).toBeVisible();
	await expect(dialog.getByLabel('Wiederholen')).toBeVisible();
	await expect(dialog.getByRole('option', { name: 'Wiederholt sich nicht' })).toBeAttached();
	// A date in the last week offers both "am vierten ..." and "am letzten ...": either will do.
	await expect(
		dialog
			.getByRole('option', {
				name: /^Jeden Monat am .* (Montag|Dienstag|Mittwoch|Donnerstag|Freitag|Samstag|Sonntag)$/u,
			})
			.first(),
	).toBeAttached();
});

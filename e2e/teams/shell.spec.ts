import { expect, test } from '@playwright/test';
import type { TeamsApp } from '../../src/ui/src/teams/app/teams-app';

test('opens settings from the more menu and changes synchronized status in the profile card', async ({
	page,
	context,
}) => {
	const room = `shell-${Date.now()}`;
	await page.goto(`/?local=1&name=Ada&room=${room}`);
	await expect(page.getByRole('heading', { name: '# General', exact: true })).toBeVisible({
		timeout: 15_000,
	});
	const more = page.getByRole('button', { name: 'Settings and more', exact: true });
	await more.focus();
	await more.press('ArrowDown');
	const settings = page.getByRole('menuitem', { name: 'Settings', exact: true });
	await expect(settings).toBeFocused();
	await settings.press('Escape');
	await expect(more).toBeFocused();
	await more.press('Enter');
	await settings.press('Enter');
	await expect(page.getByRole('dialog', { name: 'Settings', exact: true })).toBeVisible();
	await page.getByRole('button', { name: 'Close settings', exact: true }).click();
	const peer = await context.newPage();
	await peer.goto(`/?local=1&name=Bob&room=${room}`);
	await expect(peer.getByRole('heading', { name: '# General', exact: true })).toBeVisible({
		timeout: 15_000,
	});
	await page.getByRole('button', { name: 'Profile for Ada', exact: true }).click();
	const profile = page.getByRole('dialog', { name: 'Your profile', exact: true });
	const availability = profile.getByRole('combobox', { name: 'Availability', exact: true });
	await expect(availability).toBeFocused();
	await availability.selectOption('busy');
	await expect(availability).toHaveValue('busy');
	await expect
		.poll(() =>
			peer
				.locator('teams-app')
				.evaluate(
					(element) =>
						(element as TeamsApp)
							.client!.getState()
							.people.find((person) => !person.self && person.name === 'Ada')?.availability,
				),
		)
		.toBe('busy');
	await page.screenshot({ path: test.info().outputPath('profile-status.png') });
	await availability.press('Escape');
	await expect(profile).not.toBeVisible();
	await expect(page.getByRole('button', { name: 'Profile for Ada', exact: true })).toBeFocused();
	await page.setViewportSize({ width: 390, height: 844 });
	await page.getByRole('button', { name: 'Profile for Ada', exact: true }).click();
	const bounds = await profile.boundingBox();
	expect(bounds!.x).toBeGreaterThanOrEqual(0);
	expect(bounds!.x + bounds!.width).toBeLessThanOrEqual(390);
	await profile.getByRole('combobox', { name: 'Availability', exact: true }).selectOption('away');
	await expect
		.poll(() =>
			peer
				.locator('teams-app')
				.evaluate(
					(element) =>
						(element as TeamsApp)
							.client!.getState()
							.people.find((person) => !person.self && person.name === 'Ada')?.availability,
				),
		)
		.toBe('away');
	await page.screenshot({ path: test.info().outputPath('profile-mobile.png') });
	await peer.close();
});

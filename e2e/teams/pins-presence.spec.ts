import { expect, test } from '@playwright/test';
import type { TeamsApp } from '../../src/ui/src/teams/app/teams-app';

test('pins a message for everyone and shows richer presence with a status message', async ({
	browser,
}) => {
	const room = `pins-presence-${Date.now()}`;
	const context = await browser.newContext();
	const ada = await context.newPage();
	const bob = await context.newPage();
	for (const [page, name] of [
		[ada, 'Ada'],
		[bob, 'Bob'],
	] as const) {
		await page.goto(`/?local=1&name=${name}&room=${room}`);
		await expect(page.getByRole('heading', { name: '# General', exact: true })).toBeVisible({
			timeout: 15_000,
		});
	}
	await ada.getByRole('textbox', { name: 'Message', exact: true }).fill('Deploy window is Friday');
	await ada.getByRole('button', { name: 'Send message', exact: true }).click();
	const row = ada.locator('office-ui-chat-list .msg', { hasText: 'Deploy window is Friday' });
	await row.hover();
	await row.getByRole('button', { name: 'Pin', exact: true }).click();
	await expect(ada.locator('teams-pinned-bar')).toContainText('Deploy window is Friday');
	await expect(bob.locator('teams-pinned-bar')).toContainText('Deploy window is Friday');
	await expect(bob.locator('teams-pinned-bar')).toContainText('1 pinned');
	await bob.screenshot({ path: test.info().outputPath('pinned-bar.png') });
	await bob.locator('teams-pinned-bar').getByRole('button', { name: '1 pinned' }).click();
	await bob.getByRole('button', { name: /Unpin message from Ada/u }).click();
	await expect(ada.locator('teams-pinned-bar')).toHaveCount(0);

	await ada.getByRole('button', { name: 'Profile for Ada', exact: true }).click();
	const profile = ada.getByRole('dialog', { name: 'Your profile', exact: true });
	await profile.getByRole('combobox', { name: 'Reset status after' }).selectOption('1h');
	await profile.getByRole('combobox', { name: 'Availability', exact: true }).selectOption('dnd');
	await profile.getByRole('textbox', { name: 'Status message' }).fill('In a design review');
	await profile.getByRole('button', { name: 'Save message', exact: true }).click();
	await expect(profile.getByText('In a design review').first()).toBeVisible();
	await expect(profile).toContainText('Until');
	await expect
		.poll(() =>
			bob.locator('teams-app').evaluate((element) => {
				const person = (element as TeamsApp)
					.client!.getState()
					.people.find((p) => !p.self && p.name === 'Ada');
				return person ? `${person.availability}|${person.statusMessage}` : '';
			}),
		)
		.toBe('dnd|In a design review');
	await context.close();
});

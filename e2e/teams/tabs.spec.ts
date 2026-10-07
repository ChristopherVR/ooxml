import { expect, test } from '@playwright/test';
import type { TeamsApp } from '../../src/ui/src/teams/app/teams-app';

test('channel website tabs synchronize while selection and open copies stay local', async ({
	page,
	context,
}) => {
	const room = `tabs-${Date.now()}`;
	await page.goto(`/?local=1&name=Ada&room=${room}`);
	await page
		.locator('teams-app')
		.evaluate((element) => (element as TeamsApp).client!.createChannel('Project'));
	await expect(page.getByRole('heading', { name: '# Project' })).toBeVisible();
	const peer = await context.newPage();
	await peer.goto(`/?local=1&name=Bob&room=${room}`);
	await peer.locator('teams-app').evaluate((element) => ((element as TeamsApp).userId = 'bob'));
	await peer
		.locator('office-ui-channel-list')
		.getByRole('button', { name: 'Project', exact: true })
		.click();
	await expect(peer.getByRole('heading', { name: '# Project' })).toBeVisible();
	await page.getByRole('button', { name: 'Add tab', exact: true }).click();
	await page.getByRole('textbox', { name: 'Tab name', exact: true }).fill('Project site');
	await page
		.getByRole('textbox', { name: 'Tab website URL', exact: true })
		.fill('https://example.com/project');
	await page.getByRole('button', { name: 'Add website tab', exact: true }).click();
	await expect(page.getByRole('tab', { name: 'Project site', exact: true })).toHaveAttribute(
		'aria-selected',
		'true',
	);
	await expect(peer.getByRole('tab', { name: 'Project site', exact: true })).toHaveAttribute(
		'aria-selected',
		'false',
	);
	await peer.getByRole('tab', { name: 'Project site', exact: true }).click();
	await expect(peer.locator('iframe')).toHaveAttribute('src', 'https://example.com/project');
	await expect(peer.getByRole('button', { name: 'Remove tab', exact: true })).toHaveCount(0);
	page.once('dialog', (dialog) => dialog.accept('Project dashboard'));
	await page.getByRole('button', { name: 'Rename tab', exact: true }).click();
	await expect(peer.getByRole('tab', { name: 'Project dashboard', exact: true })).toBeVisible();
	page.once('dialog', (dialog) => dialog.accept());
	await page.getByRole('button', { name: 'Remove tab', exact: true }).click();
	await expect(page.getByRole('tab', { name: 'Posts', exact: true })).toHaveAttribute(
		'aria-selected',
		'true',
	);
	await expect(peer.getByRole('tab', { name: 'Project dashboard', exact: true })).toHaveCount(0);
	await expect(
		peer.getByRole('status').filter({ hasText: 'Your open copy remains here' }),
	).toBeVisible();
	await expect(peer.locator('iframe')).toHaveAttribute('src', 'https://example.com/project');
	await peer.getByRole('button', { name: 'Close preview', exact: true }).click();
	await expect(peer.getByRole('tab', { name: 'Posts', exact: true })).toHaveAttribute(
		'aria-selected',
		'true',
	);
});

import { expect, test } from '@playwright/test';
import type { TeamsApp } from '../../src/ui/src/teams/app/teams-app';

test('keeps file commands and sortable headings reachable while a long list scrolls', async ({
	page,
}) => {
	await page.goto(`/?local=1&name=Ada&room=file-layout-${Date.now()}`);
	await expect(page.getByRole('heading', { name: '# General', exact: true })).toBeVisible({
		timeout: 15_000,
	});
	await page.locator('teams-app').evaluate((element) => {
		const client = (element as TeamsApp).client!;
		client.workspace.chat.post(client.getState().selectedChannelId, {
			text: 'Project files',
			attachments: Array.from({ length: 40 }, (_, index) => ({
				name: `Notes${index + 1}.md`,
				kind: 'other' as const,
				url: `${location.origin}/notes-${index + 1}.md`,
			})),
		});
	});
	await page.getByRole('tab', { name: 'Shared', exact: true }).click();
	const panel = page.locator('teams-files-panel');
	const list = panel.locator('.file-list');
	const heading = page.getByRole('columnheader', { name: 'Name', exact: true });
	for (const width of [1280, 390]) {
		await page.setViewportSize({ width, height: 720 });
		await list.evaluate((element) => {
			element.scrollTop = element.scrollHeight;
		});
		expect(await list.evaluate((element) => element.scrollTop)).toBeGreaterThan(0);
		await expect(heading.getByRole('button')).toBeVisible();
		const header = (await heading.boundingBox())!;
		const bounds = (await list.boundingBox())!;
		const padding = await list.evaluate((element) =>
			parseFloat(getComputedStyle(element).paddingTop),
		);
		expect(header.y).toBeGreaterThanOrEqual(bounds.y);
		expect(header.y - bounds.y).toBeLessThanOrEqual(padding + 2);
		await expect(page.getByRole('button', { name: 'Upload', exact: true })).toBeVisible();
		expect(await page.evaluate(() => document.documentElement.scrollWidth)).toBeLessThanOrEqual(
			width,
		);
	}
	await heading.getByRole('button').press('Enter');
	await expect(heading).toHaveAttribute('aria-sort', 'ascending');
	await page.screenshot({ path: test.info().outputPath('files-scrolled-mobile.png') });
});

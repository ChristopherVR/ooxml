import { readFile } from 'node:fs/promises';
import { expect, test } from '@playwright/test';
import type { TeamsApp } from '../../src/ui/src/teams/app/teams-app';

// A .pptx shared in a channel opens in the reading view, drawn by the shared PowerPoint DOM
// renderer of ooxml-ui/pptx/dom (its stage is `.pptxv-stage`).
test('a .pptx attachment opens the reading view over the shared slide stage', async ({ page }) => {
	test.setTimeout(120_000);
	const errors: string[] = [];
	page.on('pageerror', (error) => errors.push(error.message));
	const bytes = await readFile(
		new URL(
			'../../src/core/pptx/__tests__/fixtures/themed-layout-placeholders.pptx',
			import.meta.url,
		),
	);
	await page.route('**/content/Algebra.pptx', (route) => route.fulfill({ body: bytes }));
	await page.goto(`/?local=1&name=Ada&room=pptx-attachment-${Date.now()}`);
	await expect(page.getByRole('heading', { name: '# General', exact: true })).toBeVisible({
		timeout: 15_000,
	});
	await page.locator('teams-app').evaluate((element) => {
		const client = (element as TeamsApp).client!;
		client.workspace.chat.post(client.getState().selectedChannelId, {
			text: 'Slides for review',
			attachments: [
				{ name: 'Algebra.pptx', kind: 'pptx', url: `${location.origin}/content/Algebra.pptx` },
			],
		});
	});
	await page.getByRole('button', { name: 'Open Algebra.pptx', exact: true }).click();
	const preview = page.locator('teams-presentation-preview');
	await expect(page.getByRole('status').filter({ hasText: 'Slide 1 of 10' })).toBeVisible({
		timeout: 90_000,
	});
	const stage = preview.locator('.pptxv-stage');
	await expect(stage).toHaveCount(1);
	await expect(stage).toHaveAttribute('aria-roledescription', 'slide');
	await expect(preview.getByRole('region', { name: 'Slide 1', exact: true })).toContainText(
		'Algebra',
	);
	expect(await stage.evaluate((node) => node.childElementCount)).toBeGreaterThan(0);
	await page.getByRole('button', { name: 'Next slide', exact: true }).click();
	await expect(preview.getByRole('region', { name: 'Slide 2', exact: true })).toBeVisible();
	await expect(stage).toHaveCount(1);
	await page.getByRole('button', { name: 'Close preview', exact: true }).click();
	await expect(preview).toHaveCount(0);
	expect(errors).toEqual([]);
});

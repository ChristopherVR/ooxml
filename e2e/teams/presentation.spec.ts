import { readFile } from 'node:fs/promises';
import { expect, test } from '@playwright/test';
import type { TeamsApp } from '../../src/ui/src/teams/app/teams-app';

test('renders PowerPoint slide images from actual bytes and navigates without a host embed', async ({
	page,
}) => {
	test.setTimeout(120_000);
	const bytes = await readFile(
		new URL(
			'../../src/core/pptx/__tests__/fixtures/themed-layout-placeholders.pptx',
			import.meta.url,
		),
	);
	await page.route('**/content/Algebra.pptx', (route) => route.fulfill({ body: bytes }));
	await page.goto('/?local=1&name=Ada&room=presentation-preview');
	await expect
		.poll(() =>
			page
				.locator('teams-app')
				.evaluate((element) => (element as TeamsApp).client?.workspace.session.roomId),
		)
		.toBe('presentation-preview');
	await expect(page.getByRole('heading', { name: '# General', exact: true })).toBeVisible();
	await page.locator('teams-app').evaluate(async (element) => {
		await (element as TeamsApp).updateComplete;
	});
	await page.locator('teams-app').evaluate((element) =>
		(element as TeamsApp).previewContent({
			attachment: { name: 'Algebra.pptx', kind: 'pptx' },
			url: `${location.origin}/content/Algebra.pptx`,
		}),
	);
	await expect(page.getByRole('status').filter({ hasText: 'Slide 1 of 10' })).toBeVisible({
		timeout: 90_000,
	});
	const image = page
		.locator('teams-presentation-preview')
		.getByRole('img', { name: 'Slide 1', exact: true });
	await expect
		.poll(() => image.evaluate((element) => (element as HTMLImageElement).naturalWidth))
		.toBeGreaterThan(0);
	const firstUrl = await image.getAttribute('src');
	expect(
		await image.evaluate(async (element) =>
			(await fetch((element as HTMLImageElement).src)).text(),
		),
	).toContain('Algebra');
	await expect(page.getByText('Static slide preview.', { exact: false })).toBeVisible();
	await expect(page.getByRole('button', { name: 'Previous slide', exact: true })).toBeDisabled();
	await page.getByRole('button', { name: 'Next slide', exact: true }).click();
	await expect(page.getByRole('status').filter({ hasText: 'Slide 2 of 10' })).toBeVisible();
	const secondImage = page
		.locator('teams-presentation-preview')
		.getByRole('img', { name: 'Slide 2', exact: true });
	await expect
		.poll(() =>
			secondImage.evaluate(async (element) =>
				(await fetch((element as HTMLImageElement).src)).text(),
			),
		)
		.toContain('stamps');
	await expect
		.poll(() => secondImage.evaluate((element) => (element as HTMLImageElement).naturalWidth))
		.toBeGreaterThan(0);
	const bounds = await secondImage.boundingBox();
	expect(bounds!.x).toBeGreaterThanOrEqual(0);
	expect(bounds!.x + bounds!.width).toBeLessThanOrEqual(page.viewportSize()!.width);
	await page.screenshot({ path: test.info().outputPath('powerpoint-preview.png') });
	const secondUrl = await secondImage.getAttribute('src');
	await page.getByText('Slide text', { exact: true }).click();
	await expect(page.locator('teams-presentation-preview pre')).toContainText(
		'Sarah collects stamps',
	);
	await page.getByRole('button', { name: 'Close preview', exact: true }).click();
	await expect(page.locator('teams-presentation-preview')).toHaveCount(0);
	for (const url of [firstUrl, secondUrl])
		expect(
			await page.evaluate(async (url) => {
				try {
					await fetch(url!);
					return false;
				} catch {
					return true;
				}
			}, url),
		).toBe(true);
});

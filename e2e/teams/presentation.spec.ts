import { readFile } from 'node:fs/promises';
import { expect, test, type Page } from '@playwright/test';
import type { TeamsApp } from '../../src/ui/src/teams/app/teams-app';

async function open(page: Page, name: string, fixture: string, room: string) {
	const bytes = await readFile(new URL(fixture, import.meta.url));
	await page.route(`**/content/${name}`, (route) => route.fulfill({ body: bytes }));
	await page.goto(`/?local=1&name=Ada&room=${room}`);
	await expect(page.getByRole('heading', { name: '# General', exact: true })).toBeVisible();
	await page
		.locator('teams-app')
		.evaluate(async (element) => await (element as TeamsApp).updateComplete);
	await page.locator('teams-app').evaluate(
		(element, name) =>
			(element as TeamsApp).previewContent({
				attachment: { name, kind: 'pptx' },
				url: `${location.origin}/content/${name}`,
			}),
		name,
	);
}

test('renders actual PowerPoint DOM content, navigates and fits compact windows', async ({
	page,
}) => {
	test.setTimeout(120_000);
	const errors: string[] = [];
	page.on('pageerror', (error) => errors.push(error.message));
	await open(
		page,
		'Algebra.pptx',
		'../../src/core/pptx/__tests__/fixtures/themed-layout-placeholders.pptx',
		'presentation-preview',
	);
	await expect(page.getByRole('status').filter({ hasText: 'Slide 1 of 10' })).toBeVisible({
		timeout: 90_000,
	});
	const preview = page.locator('teams-presentation-preview');
	await expect(preview.getByRole('region', { name: 'Slide 1', exact: true })).toContainText(
		'Algebra',
	);
	await expect(page.getByText('Reading view.', { exact: false })).toBeVisible();
	await expect(page.getByRole('button', { name: 'Previous slide', exact: true })).toBeDisabled();
	await page.getByRole('button', { name: 'Next slide', exact: true }).click();
	await expect(preview.getByRole('region', { name: 'Slide 2', exact: true })).toContainText(
		'stamps',
	);
	const frame = preview.getByRole('group', { name: 'Slide navigation canvas' });
	await frame.focus();
	await page.keyboard.press('ArrowLeft');
	await expect(page.getByRole('status').filter({ hasText: 'Slide 1 of 10' })).toBeVisible();
	await page.keyboard.press('PageDown');
	await expect(page.getByRole('status').filter({ hasText: 'Slide 2 of 10' })).toBeVisible();
	for (const width of [1280, 390]) {
		await page.setViewportSize({ width, height: 844 });
		await expect
			.poll(async () => {
				const bounds = await frame.boundingBox();
				const controls = await preview
					.getByRole('navigation', { name: 'Slide navigation' })
					.boundingBox();
				return (
					!!bounds &&
					!!controls &&
					controls.y >= 0 &&
					bounds.x >= 0 &&
					bounds.x + bounds.width <= width &&
					bounds.y >= controls.y + controls.height &&
					bounds.y + bounds.height <= 844
				);
			})
			.toBe(true);
		await expect(preview.getByRole('region', { name: 'Slide 2', exact: true })).toContainText(
			'stamps',
		);
		await page.screenshot({ path: test.info().outputPath(`powerpoint-reading-${width}.png`) });
	}
	await page.getByText('Slide text', { exact: true }).click();
	await expect(preview.locator('pre')).toContainText('Sarah collects stamps');
	await page.getByRole('button', { name: 'Close preview', exact: true }).click();
	await expect(preview).toHaveCount(0);
	expect(errors).toEqual([]);
});

test('plays embedded audio and stops and releases it when the preview closes', async ({ page }) => {
	test.setTimeout(120_000);
	const errors: string[] = [];
	page.on('pageerror', (error) => errors.push(error.message));
	await open(page, 'Audio.pptx', '../pptx/fixtures/audio-embed.pptx', 'presentation-audio');
	const preview = page.locator('teams-presentation-preview');
	const audio = preview.locator('audio');
	await expect(audio).toBeVisible({ timeout: 90_000 });
	await expect(audio).toHaveJSProperty('controls', true);
	await expect.poll(() => audio.evaluate((player) => player.duration)).toBeGreaterThan(0);
	const url = await audio.getAttribute('src');
	expect(url).toMatch(/^blob:/);
	await audio.evaluate((player) => {
		player.loop = true;
		(window as unknown as { teamsPreviewAudio: HTMLAudioElement }).teamsPreviewAudio = player;
	});
	await audio.click();
	await audio.evaluate(async (player) => await player.play());
	await expect(audio).toHaveJSProperty('paused', false);
	await page.getByRole('button', { name: 'Close preview', exact: true }).click();
	await expect(preview).toHaveCount(0);
	expect(errors).toEqual([]);
	expect(
		await page.evaluate(() => {
			const player = (window as unknown as { teamsPreviewAudio: HTMLAudioElement })
				.teamsPreviewAudio;
			return player.paused && !player.hasAttribute('src');
		}),
	).toBe(true);
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

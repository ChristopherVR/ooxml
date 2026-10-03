import { test, expect } from '@playwright/test';
import { readFileSync } from 'node:fs';
import { createHash } from 'node:crypto';
import { createVsdxFixture } from './fixture.mjs';

const legacy = readFileSync(new URL('./fixtures/owned-v11.vsd', import.meta.url));
expect(createHash('sha256').update(legacy).digest('hex')).toBe(
	'a3781bcdaadfb48f3e9669808930013197b541af164dd9e44932da6d744ef36e',
);

test('shipped worker detects legacy bytes under a modern filename and disables VSDX saving', async ({
	page,
}) => {
	await page.goto('/demo/?sample=1');
	await page
		.locator('#file')
		.setInputFiles({ name: 'legacy.vsdx', mimeType: 'application/octet-stream', buffer: legacy });
	await expect(page.locator('#edit-label')).toHaveText('LEGACY VSD PREVIEW');
	await expect(page.locator('visio-viewer [data-shape-id="7"] text')).toContainText('Hello');
	await expect(page.getByRole('button', { name: 'Download VSDX copy' })).toBeDisabled();
	await expect(page.locator('#file-state')).toContainText(
		'editing and VSDX export are unavailable',
	);
	await page.locator('#file').setInputFiles({
		name: 'modern.vsd',
		mimeType: 'application/octet-stream',
		buffer: await createVsdxFixture('Modern after legacy'),
	});
	await expect(page.locator('visio-viewer svg text')).toContainText('Modern after legacy');
	await expect(page.getByRole('button', { name: 'Download VSDX copy' })).toBeEnabled();
	await page.locator('#file').setInputFiles({
		name: 'legacy-again.vsd',
		mimeType: 'application/octet-stream',
		buffer: legacy,
	});
	await expect(page.locator('#edit-label')).toHaveText('LEGACY VSD PREVIEW');
	await expect(page.getByRole('button', { name: 'Download VSDX copy' })).toBeDisabled();
});

test('malformed binary input retains the prior bounded legacy preview', async ({ page }) => {
	await page.goto('/demo/?sample=1');
	await page
		.locator('#file')
		.setInputFiles({ name: 'owned.vsd', mimeType: 'application/octet-stream', buffer: legacy });
	await expect(page.locator('#edit-label')).toHaveText('LEGACY VSD PREVIEW');
	await page.locator('#file').setInputFiles({
		name: 'truncated.vsd',
		mimeType: 'application/octet-stream',
		buffer: legacy.subarray(0, 600),
	});
	await expect(page.locator('#error')).toBeVisible();
	await expect(page.locator('visio-viewer [data-shape-id="7"] text')).toContainText('Hello');
	await expect(page.getByRole('button', { name: 'Download VSDX copy' })).toBeDisabled();
});

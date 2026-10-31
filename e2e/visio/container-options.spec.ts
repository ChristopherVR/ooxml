import { expect, test } from '@playwright/test';
import { createVsdxFixture } from './fixture.mjs';
import { openDemo } from './demo-page';

for (const framework of ['vanilla', 'react', 'vue', 'angular', 'svelte', 'solid']) {
	test(`${framework}: Container Format locks a container, sets its resize and takes an inline width`, async ({
		page,
	}) => {
		const errors: string[] = [];
		page.on('pageerror', (error) => errors.push(error.message));
		await page.setViewportSize({ width: 1600, height: 1000 });
		await openDemo(
			page,
			framework === 'vanilla' ? '/demo/?sample=1' : `/demo-${framework}/?sample=1`,
		);
		const viewer = page.locator('visio-viewer');
		await page.locator('#file').setInputFiles({
			name: 'container-options.vsdx',
			mimeType: 'application/vnd.ms-visio.drawing',
			buffer: await createVsdxFixture('Existing shape'),
		});
		await expect(viewer.locator('svg.paper')).toContainText('Existing shape');
		const status = viewer.locator('[data-status]');

		await viewer.getByRole('tab', { name: 'Insert', exact: true }).click();
		await viewer.locator('office-ui-gallery[command="container"] .trigger').click();
		await viewer.locator('[data-gallery-popup="container"] [data-gallery-item="plain"]').click();
		await expect(status).toHaveText('Inserted a Plain container.');

		const tab = viewer.getByRole('tab', { name: 'Container Format', exact: true });
		await expect(tab).toBeVisible();
		await tab.click();
		const panel = viewer.locator('[data-ribbon-tab="container-format"]');

		await panel.locator('[command="container-lock"] button').click();
		await expect(status).toHaveText('Locked the container.');
		await panel.locator('[command="container-lock"] button').click();
		await expect(status).toHaveText('Unlocked the container.');

		const resize = panel.locator('office-ui-menu-button[data-menu="container-resize"]');
		await resize.getByRole('button', { name: 'Automatic Resize', exact: true }).click();
		await resize
			.getByRole('menuitemcheckbox', { name: 'Always fit to contents', exact: true })
			.click();
		await expect(status).toHaveText('Changed the automatic resize of the container.');

		// The inline Width box resizes the selected container in one step.
		const width = panel.locator('[data-inline-size="width"]');
		await width.fill('5');
		await width.press('Enter');
		await expect(width).toHaveValue(/^5(\.0+)?$/);
		await width.press('Tab');
		await expect(width).toHaveValue(/^5(\.0+)?$/);
		expect(errors).toEqual([]);
	});
}

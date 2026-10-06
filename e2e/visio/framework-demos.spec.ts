import { test, expect } from '@playwright/test';
import { createVsdxFixture } from './fixture.mjs';
import { fileBackstage } from './ribbon.js';

// GitHub Pages serves one demo per framework; each mounts the viewer through its own binding and
// the shared demo workspace, so all of them must behave the same.
for (const framework of ['react', 'vue', 'angular', 'svelte', 'solid']) {
	test(`the ${framework} demo opens, edits and switches drawings through its binding`, async ({
		page,
	}) => {
		const errors: string[] = [];
		page.on('pageerror', (error) => errors.push(error.message));
		await page.setViewportSize({ width: 1440, height: 900 });
		await page.goto(`/demo-${framework}/?sample=1`);
		const viewer = page.locator('visio-viewer');
		await expect(viewer.locator('svg.paper')).toHaveAttribute('aria-label', 'Release workflow');
		await expect(viewer.getByRole('tab', { name: 'Home', exact: true })).toBeVisible();
		await expect(page.locator('#note-count')).not.toHaveText('');

		// A file opened from the workspace loads through the binding's handle.
		await page.locator('#file').setInputFiles({
			name: `${framework}.vsdx`,
			mimeType: 'application/vnd.ms-visio.drawing',
			buffer: await createVsdxFixture(`${framework} shape`),
		});
		await expect(page.locator('#file-name')).toHaveText(`${framework}.vsdx`);
		await expect(viewer.locator('svg.paper')).toContainText(`${framework} shape`);
		await expect(page.locator('#edit-label')).toHaveText('ORIGINAL');

		// Selecting a shape reaches the workspace through the binding's events.
		await viewer.locator('svg.paper [data-shape-id]').first().click();
		await expect(page.locator('#selection')).toContainText('Shape ID');

		// The sample template goes back through the framework's own document property.
		await fileBackstage(viewer, 'new');
		await viewer.locator('.template-card').click();
		await expect(viewer.locator('svg.paper')).toHaveAttribute('aria-label', 'Release workflow');
		await expect(page.locator('#file-name')).toHaveText('Sample workflow');
		expect(errors).toEqual([]);
	});
}

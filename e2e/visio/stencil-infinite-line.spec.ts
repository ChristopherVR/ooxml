import { readFileSync } from 'node:fs';
import { expect, test } from '@playwright/test';
import { openDemo } from './demo-page';

for (const framework of ['vanilla', 'react', 'vue', 'angular', 'svelte', 'solid']) {
	test(`${framework}: native infinite-line masters have drawn stencil previews`, async ({
		page,
	}) => {
		await page.setViewportSize({ width: 1440, height: 1000 });
		await openDemo(
			page,
			framework === 'vanilla' ? '/demo/?sample=1' : `/demo-${framework}/?sample=1`,
		);
		await page.locator('#file').setInputFiles({
			name: 'infinite-line-masters.vsdx',
			mimeType: 'application/vnd.ms-visio.drawing',
			buffer: readFileSync(
				new URL('./fixtures/infinite-line-masters-native.vsdx', import.meta.url),
			),
		});
		await expect(page.locator('#file-name')).toHaveText('infinite-line-masters.vsdx');
		const stencil = page.locator('visio-viewer [data-stencil="document"]');
		for (const [name, path] of [
			['Diagonal', 'M 0 0 L 3 3'],
			['Horizontal', 'M 0 0.5 L 4 0.5'],
			['Vertical', 'M 0.5 0 L 0.5 3'],
		]) {
			const preview = stencil
				.getByRole('button', { name: name!, exact: true })
				.locator('svg.master-preview');
			await expect(preview).toBeVisible();
			const line = preview.locator(`path[d="${path}"]`).first();
			await expect(line).toBeAttached();
			expect(
				await line.evaluate((node) => (node as SVGPathElement).getTotalLength()),
			).toBeGreaterThan(0);
		}
		if (framework === 'vanilla')
			await stencil.screenshot({ path: test.info().outputPath('infinite-line-previews.png') });
	});
}

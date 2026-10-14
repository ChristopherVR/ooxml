import { expect, test } from '@playwright/test';
import { openDemo } from './demo-page';

for (const framework of ['vanilla', 'react', 'vue', 'angular', 'svelte', 'solid']) {
	test(`${framework}: Visio keeps tab height and supports native ribbon collapse`, async ({
		page,
	}) => {
		await page.setViewportSize({ width: 2200, height: 900 });
		await openDemo(page, `/demo-${framework}/?sample=1`);
		const ribbon = page.locator('visio-viewer').locator('office-ui-ribbon');
		const names = await ribbon.getByRole('tab').allTextContents();
		expect(names).toEqual([
			'Home',
			'Insert',
			'Design',
			'Data',
			'Process',
			'Review',
			'View',
			'Help',
		]);
		const height = (await ribbon.boundingBox())!.height;
		for (const name of names) {
			const tab = ribbon.getByRole('tab', { name, exact: true });
			await tab.click();
			await expect(tab).toHaveAttribute('aria-selected', 'true');
			expect((await ribbon.boundingBox())!.height).toBe(height);
		}
		const home = ribbon.getByRole('tab', { name: 'Home', exact: true });
		await home.focus();
		await page.keyboard.press('Control+F1');
		await expect(ribbon).toHaveAttribute('collapsed', '');
		await home.dblclick();
		await expect(ribbon).not.toHaveAttribute('collapsed', '');
	});
}

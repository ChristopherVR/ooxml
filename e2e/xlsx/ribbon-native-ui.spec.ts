import { expect, test } from '@playwright/test';
import JSZip from 'jszip';
import { editor, FRAMEWORKS, newWorkbook, openLanding } from './helpers';

for (const framework of FRAMEWORKS) {
	test(`${framework}: Excel Home labels and responsive Cell Styles`, async ({ page }) => {
		await page.setViewportSize({ width: 2200, height: 900 });
		await openLanding(page, framework);
		await newWorkbook(page, framework);
		const root = editor(page);
		for (const label of ['Cut', 'Copy', 'Format Painter', 'Wrap Text', 'Merge & Center', 'AutoSum'])
			await expect(root.getByRole('button', { name: label, exact: true })).toContainText(label);
		const gallery = root.locator('.xve-cell-styles');
		await expect(gallery).toHaveAttribute('mode', 'inline');
		const good = gallery.locator('.strip .tile[aria-label="Good"]');
		await expect(good).toBeVisible();
		await good.click();
		const bytes = await root.evaluate(async (node) =>
			Array.from(await (node as HTMLElement & { saveBytes(): Promise<Uint8Array> }).saveBytes()),
		);
		const zip = await JSZip.loadAsync(Uint8Array.from(bytes));
		expect(await zip.file('xl/styles.xml')!.async('string')).toMatch(/C6EFCE/i);
		await page.setViewportSize({ width: 1400, height: 900 });
		await expect(gallery).not.toHaveAttribute('mode', 'inline');
		await expect(gallery.getByRole('button', { name: 'Cell Styles', exact: true })).toBeVisible();
	});
}

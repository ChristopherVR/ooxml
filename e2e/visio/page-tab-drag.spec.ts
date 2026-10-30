import { test, expect, type Locator, type Page } from '@playwright/test';
import { createVsdxFixture } from './fixture.mjs';
import { downloadCopy } from './ribbon';
import { openDemo } from './demo-page';

/** Press on the centre of `from`, move across the bar and release at `x` (page coordinates). */
async function drag(page: Page, from: Locator, x: number, release = true): Promise<void> {
	const box = (await from.boundingBox())!;
	const y = box.y + box.height / 2;
	await page.mouse.move(box.x + box.width / 2, y);
	await page.mouse.down();
	await page.mouse.move(box.x + box.width / 2 + 8, y, { steps: 2 });
	await page.mouse.move(x, y, { steps: 6 });
	if (release) await page.mouse.up();
}

for (const framework of ['vanilla', 'react', 'vue', 'angular', 'svelte', 'solid']) {
	test(`${framework}: drags page tabs to reorder pages, cancels with Escape, undoes and keeps the order when saved`, async ({
		page,
	}) => {
		await page.setViewportSize({ width: 1440, height: 900 });
		await openDemo(
			page,
			framework === 'vanilla' ? '/demo/?sample=1' : `/demo-${framework}/?sample=1`,
		);
		const viewer = page.locator('visio-viewer');
		const tabs = viewer.locator('.page-tabs');
		const tab = (name: string) => tabs.getByRole('tab', { name, exact: true });
		await page.locator('#file').setInputFiles({
			name: 'pages.vsdx',
			mimeType: 'application/vnd.ms-visio.drawing',
			buffer: await createVsdxFixture('Original page shape'),
		});
		await expect(page.locator('#file-name')).toHaveText('pages.vsdx');
		const insert = tabs.getByRole('button', { name: 'Insert Page', exact: true });
		await insert.click();
		await insert.click();
		await expect(tabs.getByRole('tab')).toHaveCount(3);
		const names = () => tabs.getByRole('tab').allTextContents();
		const original = await names();
		const [first, second, third] = original as [string, string, string];

		// A click without movement still switches pages.
		await tab(first).click();
		await expect(tab(first)).toHaveAttribute('aria-selected', 'true');

		// Escape abandons the drag and leaves the order alone.
		const target = (await tab(third).boundingBox())!;
		await drag(page, tab(first), target.x + target.width - 4, false);
		await expect(tabs.locator('.marker')).toHaveCount(1);
		await page.keyboard.press('Escape');
		await page.mouse.up();
		await expect(tabs.locator('.marker')).toHaveCount(0);
		expect(await names()).toEqual(original);

		// Dropping on the right edge of the last tab moves the page to the end.
		await drag(page, tab(first), target.x + target.width - 4);
		await expect.poll(names).toEqual([second, third, first]);

		// One undo step restores the order; redo applies it again.
		await viewer.locator('.qat [data-command="undo"]').click();
		await expect.poll(names).toEqual(original);
		await viewer.locator('.qat [data-command="redo"]').click();
		await expect.poll(names).toEqual([second, third, first]);

		// Save, reopen: the order persists.
		const downloadButton = await downloadCopy(viewer);
		const pending = page.waitForEvent('download');
		await downloadButton.click();
		const stream = await (await pending).createReadStream();
		const chunks: Buffer[] = [];
		for await (const chunk of stream!) chunks.push(Buffer.from(chunk));
		await page.locator('#file').setInputFiles({
			name: 'reordered-pages.vsdx',
			mimeType: 'application/vnd.ms-visio.drawing',
			buffer: Buffer.concat(chunks),
		});
		await expect(page.locator('#file-name')).toHaveText('reordered-pages.vsdx');
		await expect.poll(names).toEqual([second, third, first]);
	});
}

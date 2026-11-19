import { expect, test, type Page } from '@playwright/test';
import { openSample } from './helpers';

/** The share of a PNG's pixels that match each system colour, decoded in the page. */
async function systemColourShares(page: Page, png: Buffer, names: string[]) {
	return page.evaluate(
		async ([base64, names]) => {
			const probe = document.createElement('i');
			probe.style.forcedColorAdjust = 'none';
			document.body.append(probe);
			const targets = (names as string[]).map((name) => {
				probe.style.color = name;
				return getComputedStyle(probe).color.match(/\d+/g)!.slice(0, 3).map(Number);
			});
			probe.remove();
			const bitmap = await createImageBitmap(
				await (await fetch(`data:image/png;base64,${base64}`)).blob(),
			);
			const canvas = new OffscreenCanvas(bitmap.width, bitmap.height);
			const context = canvas.getContext('2d')!;
			context.drawImage(bitmap, 0, 0);
			const data = context.getImageData(0, 0, bitmap.width, bitmap.height).data;
			const counts = targets.map(() => 0);
			for (let i = 0; i < data.length; i += 4)
				targets.forEach(([r, g, b], index) => {
					if (
						Math.abs(data[i]! - r!) + Math.abs(data[i + 1]! - g!) + Math.abs(data[i + 2]! - b!) <
						24
					)
						counts[index]!++;
				});
			return counts.map((count) => count / (data.length / 4));
		},
		[png.toString('base64'), names] as const,
	);
}

for (const scheme of ['light', 'dark'] as const) {
	test(`forced colours (${scheme} palette): pressed buttons stay filled and the active sheet stays marked`, async ({
		page,
	}) => {
		await page.emulateMedia({ forcedColors: 'active', colorScheme: scheme });
		await openSample(page);
		const wrap = page.getByRole('button', { name: 'Wrap Text' }).first();
		await wrap.click();
		if ((await wrap.getAttribute('aria-pressed')) !== 'true') await wrap.click();
		await expect(wrap).toHaveAttribute('aria-pressed', 'true');
		await page.mouse.move(0, 0);
		const [fill = 0, glyph = 0] = await systemColourShares(page, await wrap.screenshot(), [
			'Highlight',
			'HighlightText',
		]);
		expect(fill, 'pressed button is filled with Highlight').toBeGreaterThan(0.3);
		expect(fill + glyph, 'its icon is HighlightText on Highlight').toBeGreaterThan(0.8);

		const tab = page.locator('xlsx-editor .xst-tab.xst-active');
		const [selected] = await systemColourShares(page, await tab.screenshot(), ['Highlight']);
		expect(selected, 'active sheet tab is filled with Highlight').toBeGreaterThan(0.5);
	});
}

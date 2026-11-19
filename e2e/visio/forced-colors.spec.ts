import { expect, test, type Page } from '@playwright/test';
import { openDemo } from './demo-page';

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
			const blob = await (await fetch(`data:image/png;base64,${base64}`)).blob();
			const bitmap = await createImageBitmap(blob);
			const context = new OffscreenCanvas(bitmap.width, bitmap.height).getContext('2d')!;
			context.drawImage(bitmap, 0, 0);
			const data = context.getImageData(0, 0, bitmap.width, bitmap.height).data;
			const counts = targets.map(() => 0);
			for (let i = 0; i < data.length; i += 4)
				targets.forEach(([r, g, b], index) => {
					const distance =
						Math.abs(data[i]! - r!) + Math.abs(data[i + 1]! - g!) + Math.abs(data[i + 2]! - b!);
					if (distance < 24) counts[index]!++;
				});
			return counts.map((count) => count / (data.length / 4));
		},
		[png.toString('base64'), names] as const,
	);
}

for (const scheme of ['light', 'dark'] as const) {
	test(`forced colours (${scheme} palette): the pressed Pointer Tool keeps its label`, async ({
		page,
	}) => {
		await page.emulateMedia({ forcedColors: 'active', colorScheme: scheme });
		await page.setViewportSize({ width: 1440, height: 1000 });
		await openDemo(page);
		const pointer = page
			.locator('visio-viewer .toolbar')
			.getByRole('button', { name: 'Pointer Tool', exact: true });
		await expect(pointer).toHaveAttribute('aria-pressed', 'true');
		await page.mouse.move(0, 0);
		// The label sits on Highlight with HighlightText glyphs. With the browser's text backplate
		// the label box was painted Canvas, which hid the HighlightText label in dark palettes.
		const [highlight, text] = await systemColourShares(
			page,
			await pointer.locator('.label').screenshot(),
			['Highlight', 'HighlightText'],
		);
		expect(highlight, 'label background is Highlight').toBeGreaterThan(0.3);
		expect(text, 'label glyphs are HighlightText').toBeGreaterThan(0.02);
	});
}

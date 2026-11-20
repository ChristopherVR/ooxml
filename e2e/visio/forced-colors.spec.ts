import { expect, test, type Page } from '@playwright/test';
import { openDemo } from './demo-page';

/**
 * The share of a PNG's pixels that match each system colour, decoded in the page. A system colour
 * may be translucent (Chromium on Linux gives Highlight 80% alpha), so each is matched as it
 * paints over Canvas, the backdrop forced colours give the page: one layer, or two where a border
 * lies over a background of the same colour.
 */
async function systemColourShares(page: Page, png: Buffer, names: string[]) {
	return page.evaluate(
		async ([base64, names]) => {
			const probe = document.createElement('i');
			probe.style.forcedColorAdjust = 'none';
			document.body.append(probe);
			const rgba = (name: string) => {
				probe.style.color = name;
				const [r = 0, g = 0, b = 0, a = 1] = getComputedStyle(probe)
					.color.match(/[\d.]+/g)!
					.map(Number);
				return { rgb: [r, g, b], a };
			};
			const backdrop = rgba('Canvas').rgb;
			const shades = (names as string[]).map((name) => {
				const { rgb, a } = rgba(name);
				const over = (under: number[]) =>
					rgb.map((channel, i) => channel * a + under[i]! * (1 - a));
				return [over(backdrop), over(over(backdrop))];
			});
			probe.remove();
			const blob = await (await fetch(`data:image/png;base64,${base64}`)).blob();
			const bitmap = await createImageBitmap(blob);
			const context = new OffscreenCanvas(bitmap.width, bitmap.height).getContext('2d')!;
			context.drawImage(bitmap, 0, 0);
			const data = context.getImageData(0, 0, bitmap.width, bitmap.height).data;
			const counts = shades.map(() => 0);
			for (let i = 0; i < data.length; i += 4)
				shades.forEach((colour, index) => {
					const near = ([r, g, b]: number[]) =>
						Math.abs(data[i]! - r!) + Math.abs(data[i + 1]! - g!) + Math.abs(data[i + 2]! - b!) <
						24;
					if (colour.some(near)) counts[index]!++;
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

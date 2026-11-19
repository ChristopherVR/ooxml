import { expect, test, type Page } from '@playwright/test';
import { editor, openSample, sheetTabs } from './helpers';

/**
 * Themes recolour the chrome, never the document. The sheet's cells, gridlines and paper must be
 * pixel-identical in light, dark and a sentinel theme that sets every chrome colour token to a
 * garish value; only the selection (hidden here) may follow the accent.
 */
const CANVAS_TOKENS = ['--office-paper', '--office-canvas-gridline', '--office-canvas-link'];

async function applyTheme(page: Page, theme: 'light' | 'dark' | 'sentinel') {
	await page.evaluate(
		({ theme, canvas }) => {
			const host = document.querySelector('xlsx-editor') as HTMLElement;
			const root = host.shadowRoot!;
			if (!root.getElementById('hide-selection')) {
				const style = document.createElement('style');
				style.id = 'hide-selection';
				style.textContent =
					'.xg-overlay,.xg-remote-layer{visibility:hidden!important}.xg-scroller{scrollbar-width:none}';
				root.append(style);
			}
			for (const name of [...host.style])
				if (name.startsWith('--office-')) host.style.removeProperty(name);
			host.setAttribute('data-office-theme', theme === 'dark' ? 'dark' : 'light');
			if (theme !== 'sentinel') return;
			// Every colour token the page declares, except the document canvas, set to a sentinel.
			const declared = getComputedStyle(document.documentElement);
			for (const name of declared) {
				if (!name.startsWith('--office-') || canvas.includes(name)) continue;
				const value = declared.getPropertyValue(name).trim();
				if (/^(#|rgb|hsl|color-mix|var\()/.test(value))
					host.style.setProperty(name, name.includes('foreground') ? '#00ff00' : '#ff00ff');
			}
		},
		{ theme, canvas: CANVAS_TOKENS },
	);
	// Let the attribute and token changes paint.
	await page.evaluate(
		() => new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve))),
	);
}

/** The cell area of the grid (below and right of the headers), as a PNG. */
async function cellArea(page: Page): Promise<Buffer> {
	const clip = await page.evaluate(() => {
		const root = document.querySelector('xlsx-editor')!.shadowRoot!;
		const view = root.querySelector('.xg-view')!.getBoundingClientRect();
		const boxes = [...root.querySelectorAll<HTMLElement>('.xg-hbox')]
			.filter((box) => box.offsetParent)
			.map((box) => box.getBoundingClientRect());
		const left = Math.max(
			view.left,
			...boxes.filter((b) => b.height > b.width).map((b) => b.right),
		);
		const top = Math.max(
			view.top,
			...boxes.filter((b) => b.width >= b.height).map((b) => b.bottom),
		);
		return { x: left, y: top, width: view.right - left - 20, height: view.bottom - top - 20 };
	});
	expect(clip.width).toBeGreaterThan(200);
	expect(clip.height).toBeGreaterThan(100);
	return page.screenshot({ clip, animations: 'disabled', caret: 'hide' });
}

/** How many pixels differ between two same-sized PNGs, decoded in the page. */
async function changedPixels(page: Page, a: Buffer, b: Buffer): Promise<number> {
	return page.evaluate(
		async ([first, second]) => {
			const decode = async (base64: string) => {
				const bitmap = await createImageBitmap(
					await (await fetch(`data:image/png;base64,${base64}`)).blob(),
				);
				const canvas = new OffscreenCanvas(bitmap.width, bitmap.height);
				const context = canvas.getContext('2d')!;
				context.drawImage(bitmap, 0, 0);
				return context.getImageData(0, 0, bitmap.width, bitmap.height).data;
			};
			const [x, y] = await Promise.all([decode(first!), decode(second!)]);
			if (x.length !== y.length) return Number.POSITIVE_INFINITY;
			let changed = 0;
			for (let i = 0; i < x.length; i += 4)
				if (x[i] !== y[i] || x[i + 1] !== y[i + 1] || x[i + 2] !== y[i + 2]) changed++;
			return changed;
		},
		[a.toString('base64'), b.toString('base64')],
	);
}

test('themes never change the sheet: 0 changed cell pixels in light, dark and a sentinel theme', async ({
	page,
}) => {
	await openSample(page);
	await expect(editor(page)).toBeVisible();
	const tabBar = () => sheetTabs(page).evaluate((el) => getComputedStyle(el).backgroundColor);
	await applyTheme(page, 'light');
	const light = await cellArea(page);
	const lightTabBar = await tabBar();
	await applyTheme(page, 'dark');
	const dark = await cellArea(page);
	await applyTheme(page, 'sentinel');
	const sentinel = await cellArea(page);
	expect(await changedPixels(page, light, dark)).toBe(0);
	expect(await changedPixels(page, light, sentinel)).toBe(0);
	// The sentinel really did recolour the chrome around the sheet.
	expect(await tabBar()).not.toBe(lightTabBar);
});

test('the active sheet tab reads at 4.5:1 in light and dark', async ({ page }) => {
	await openSample(page);
	for (const theme of ['light', 'dark'] as const) {
		await applyTheme(page, theme);
		const [fg, bg] = await page.evaluate(() => {
			const tab = document
				.querySelector('xlsx-editor')!
				.shadowRoot!.querySelector<HTMLElement>('.xst-tab.xst-active')!;
			const style = getComputedStyle(tab);
			// Resolve any colour syntax (color-mix() included) to 8-bit sRGB through a canvas.
			const context = new OffscreenCanvas(1, 1).getContext('2d')!;
			const rgb = (value: string) => {
				context.clearRect(0, 0, 1, 1);
				context.fillStyle = value;
				context.fillRect(0, 0, 1, 1);
				return [...context.getImageData(0, 0, 1, 1).data.slice(0, 3)];
			};
			return [rgb(style.color), rgb(style.backgroundColor)];
		});
		const luminance = (value: number[]) => {
			const [r, g, b] = value.map((c) => {
				const s = c / 255;
				return s <= 0.04045 ? s / 12.92 : ((s + 0.055) / 1.055) ** 2.4;
			});
			return 0.2126 * r! + 0.7152 * g! + 0.0722 * b!;
		};
		const [a, b] = [luminance(fg), luminance(bg)];
		expect(
			(Math.max(a, b) + 0.05) / (Math.min(a, b) + 0.05),
			`${theme}: ${fg} on ${bg}`,
		).toBeGreaterThanOrEqual(4.5);
	}
});

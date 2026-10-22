import { expect, test } from '@playwright/test';
import { createVsdxFixture } from './fixture.mjs';
import type { ViewerController } from 'ooxml-ui/visio';
declare global {
	interface Window {
		nativeBlocks: { controller: ViewerController };
	}
}
for (const framework of ['react', 'vue', 'angular', 'svelte', 'solid', 'vanilla'])
	test(`${framework}: shared theme reaches native controls and the shadow canvas`, async ({
		page,
	}) => {
		await page.emulateMedia({ colorScheme: 'light' });
		await page.goto(`/native-blocks.html?framework=${framework}`);
		await expect(page.locator('output')).toHaveText('view:1');
		await page.locator('#content').setInputFiles({
			name: 'theme.vsdx',
			mimeType: 'application/octet-stream',
			buffer: await createVsdxFixture('Theme preserves drawing'),
		});
		await expect(page.locator('svg.paper text')).toContainText('Theme preserves drawing');
		await page.locator('svg.paper [data-shape-id="1"]').click();
		const before = await page.evaluate(() =>
			Array.from(window.nativeBlocks.controller.exportVsdx().bytes),
		);
		const read = () =>
			page.evaluate(() => {
				const toolbar = document.querySelector<HTMLElement>('[data-office-block="toolbar"]')!;
				const tabs = document.querySelector<HTMLElement>('[data-office-block="tabs"]')!;
				const inspector = document.querySelector<HTMLElement>('[data-office-block="inspector"]')!;
				const canvasHost = Array.from(document.querySelectorAll<HTMLElement>('#app *')).find(
					(host) => host.shadowRoot?.querySelector('[part="canvas"]'),
				)!;
				const viewport = canvasHost.shadowRoot!.querySelector<HTMLElement>('[part="canvas"]')!;
				const button = toolbar.querySelector<HTMLButtonElement>('button:last-child')!;
				const disabled = toolbar.querySelector<HTMLButtonElement>('button:disabled')!;
				const color = (element: Element) => {
					const style = getComputedStyle(element);
					return {
						background: style.backgroundColor,
						color: style.color,
						border: style.borderTopColor,
						font: style.fontFamily,
					};
				};
				return {
					toolbar: color(toolbar),
					tabs: color(tabs),
					inspector: color(inspector),
					canvas: color(viewport),
					button: color(button),
					disabledOpacity: Number(getComputedStyle(disabled).opacity),
					paper: getComputedStyle(canvasHost.shadowRoot!.querySelector('svg.paper')!)
						.backgroundColor,
					selected: window.nativeBlocks.controller.state.selectedShapes.length,
				};
			});
		const initial = await read();
		expect(initial.selected).toBe(1);
		expect(initial.toolbar.color).toBe(initial.tabs.color);
		expect(initial.inspector.color).toBe(initial.toolbar.color);
		expect(initial.canvas.color).toBe(initial.toolbar.color);
		await page.evaluate(() => (document.documentElement.dataset.officeTheme = 'dark'));
		const dark = await read();
		expect(dark.toolbar.background).not.toBe(initial.toolbar.background);
		expect(dark.toolbar.color).not.toBe(initial.toolbar.color);
		expect(dark.canvas.background).not.toBe(initial.canvas.background);
		await page.evaluate(
			() => (document.querySelector<HTMLElement>('#app')!.dataset.officeTheme = 'light'),
		);
		const scopedLight = await read();
		expect(scopedLight.toolbar.color).toBe(initial.toolbar.color);
		expect(scopedLight.canvas.background).toBe(initial.canvas.background);
		await page.evaluate(
			() => (document.querySelector<HTMLElement>('#app')!.dataset.officeTheme = 'dark'),
		);
		const scopedDark = await read();
		expect(scopedDark.toolbar.color).toBe(dark.toolbar.color);
		expect(scopedDark.canvas.background).toBe(dark.canvas.background);
		await page.evaluate(() => {
			document.documentElement.dataset.officeTheme = 'light';
			document.querySelector<HTMLElement>('#app')!.dataset.officeTheme = 'auto';
		});
		await page.emulateMedia({ colorScheme: 'dark' });
		const autoDark = await read();
		expect(autoDark.toolbar.color).toBe(dark.toolbar.color);
		await page.emulateMedia({ colorScheme: 'light' });
		const autoLight = await read();
		expect(autoLight.toolbar.color).toBe(initial.toolbar.color);
		await page.evaluate(() => {
			const host = document.querySelector<HTMLElement>('#app')!;
			for (const [key, value] of Object.entries({
				'--office-background': '#243447',
				'--office-surface': '#243447',
				'--office-foreground': '#e7effb',
				'--office-border': '#617992',
				'--office-accent': '#a78bfa',
				'--office-ring': '#c4b5fd',
				'--office-font': 'Courier New, monospace',
			}))
				host.style.setProperty(key, value);
		});
		const custom = await read();
		for (const block of Object.values({
			toolbar: custom.toolbar,
			tabs: custom.tabs,
			inspector: custom.inspector,
			canvas: custom.canvas,
			button: custom.button,
		})) {
			expect(block.background).toBe('rgb(36, 52, 71)');
			expect(block.color).toBe('rgb(231, 239, 251)');
			expect(block.font, JSON.stringify(custom)).toContain('Courier New');
		}
		expect(custom.tabs.border).toBe('rgb(97, 121, 146)');
		expect(custom.disabledOpacity).toBeLessThan(1);
		const fit = page.getByRole('button', { name: 'Fit to Window', exact: true });
		await page.keyboard.press('Tab');
		await fit.focus();
		expect(await fit.evaluate((button) => getComputedStyle(button).outlineColor)).toBe(
			'rgb(196, 181, 253)',
		);
		await page.evaluate(() => {
			const host = document.querySelector<HTMLElement>('#app')!;
			host.style.setProperty('--vv-surface', '#3e2447');
			host.style.setProperty('--vv-background', '#3e2447');
			host.style.setProperty('--vv-ink', '#f5dffb');
			host.style.setProperty('--vv-border', '#a877b4');
			host.style.setProperty('--vv-font', 'Georgia, serif');
		});
		const legacy = await read();
		for (const block of [
			legacy.toolbar,
			legacy.tabs,
			legacy.inspector,
			legacy.canvas,
			legacy.button,
		]) {
			expect(block.background).toBe('rgb(62, 36, 71)');
			expect(block.color).toBe('rgb(245, 223, 251)');
			expect(block.font).toContain('Georgia');
		}
		expect(legacy.tabs.border).toBe('rgb(168, 119, 180)');
		expect(legacy.paper).toBe('rgb(255, 255, 255)');
		expect(legacy.selected).toBe(1);
		const after = await page.evaluate(() =>
			Array.from(window.nativeBlocks.controller.exportVsdx().bytes),
		);
		expect(after).toEqual(before);
	});

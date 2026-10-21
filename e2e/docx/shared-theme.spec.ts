import { expect, test } from '@playwright/test';
import { openSample } from './helpers';
import type { NativeHarness } from '../../demos/docx/demo-vanilla/building-blocks-types';
for (const framework of ['react', 'vue', 'angular', 'svelte', 'solid']) {
	test(`${framework}: shared Office theme scopes style native controls and Word chrome reactively`, async ({
		page,
	}) => {
		await page.goto(`/building-blocks.html?framework=${framework}`);
		await expect(page.locator('[data-office-block="toolbar"]')).toBeVisible();
		const snapshot = () =>
			page.evaluate(() => {
				const root = document.querySelector<HTMLElement>('#native-host')!;
				const editor = root.querySelector('docx-editor')!;
				const chrome = editor.shadowRoot!;
				const toolbar = root.querySelector<HTMLElement>('[data-office-block="toolbar"]')!;
				const probe = document.createElement('span');
				root.append(probe);
				const token = (name: string) => {
					probe.style.color = `var(${name})`;
					return getComputedStyle(probe).color;
				};
				const result = {
					canvas: getComputedStyle(chrome.querySelector('.dve-canvas')!).backgroundColor,
					paper: getComputedStyle(chrome.querySelector('.dve-paper')!).backgroundColor,
					toolbar: getComputedStyle(toolbar).backgroundColor,
					buttonText: getComputedStyle(toolbar.querySelector('button')!).color,
					buttonFont: getComputedStyle(toolbar.querySelector('button')!).fontFamily,
					buttonSize: getComputedStyle(toolbar.querySelector('button')!).fontSize,
					buttonRadius: getComputedStyle(toolbar.querySelector('button')!).borderRadius,
					editorFont: getComputedStyle(editor).fontFamily,
					background: token('--office-background'),
					surface: token('--office-surface'),
					foreground: token('--office-foreground'),
				};
				probe.remove();
				return result;
			});
		const model = await page.evaluate(() => {
			const harness: NativeHarness = window.nativeWord!;
			return JSON.stringify(harness.handle.element.documentModel);
		});
		await page.evaluate(() =>
			document.querySelector('#native-host')!.setAttribute('data-office-theme', 'light'),
		);
		const light = await snapshot();
		expect(light.canvas).toBe(light.surface);
		expect(light.paper).toBe('rgb(255, 255, 255)');
		expect(light.buttonText).toBe(light.foreground);
		await page.evaluate(() =>
			document.querySelector('#native-host')!.setAttribute('data-office-theme', 'dark'),
		);
		const dark = await snapshot();
		expect(dark.canvas).toBe(dark.surface);
		expect(dark.canvas).not.toBe(light.canvas);
		expect(dark.toolbar).not.toBe(light.toolbar);
		expect(dark.buttonText).toBe(dark.foreground);
		await page.evaluate(() => {
			const root = document.querySelector<HTMLElement>('#native-host')!;
			root.style.setProperty('--office-surface', 'rgb(17, 34, 51)');
			root.style.setProperty('--office-background', 'rgb(23, 45, 67)');
			root.style.setProperty('--office-foreground', 'rgb(221, 222, 223)');
			root.style.setProperty('--office-font', 'monospace');
			root.style.setProperty('--office-font-size', '17px');
			root.style.setProperty('--office-radius', '9px');
			root.style.setProperty('--office-ring', 'rgb(1, 2, 3)');
		});
		const custom = await snapshot();
		expect(custom.canvas).toBe('rgb(17, 34, 51)');
		expect(custom.buttonText).toBe('rgb(221, 222, 223)');
		expect(custom.buttonFont).toBe('monospace');
		expect(custom.editorFont).toBe('monospace');
		expect(custom.buttonSize).toBe('17px');
		expect(custom.buttonRadius).toBe('9px');
		const bold = page.getByRole('button', { name: 'Bold', exact: true });
		await bold.focus();
		await page.keyboard.press('Tab');
		await page.keyboard.press('Shift+Tab');
		expect(await bold.evaluate((node) => getComputedStyle(node).outlineColor)).toBe('rgb(1, 2, 3)');
		expect([custom.background, custom.surface]).toContain(custom.toolbar);
		await page.evaluate(() => {
			const root = document.querySelector<HTMLElement>('#native-host')!;
			root.style.setProperty('--dve-muted', 'rgb(51, 41, 31)');
			root.style.setProperty('--dve-card', 'rgb(64, 54, 44)');
			root.style.setProperty('--dve-foreground', 'rgb(230, 220, 210)');
		});
		const legacy = await snapshot();
		expect(legacy.canvas).toBe('rgb(51, 41, 31)');
		expect(legacy.toolbar).toBe('rgb(64, 54, 44)');
		expect(legacy.buttonText).toBe('rgb(230, 220, 210)');
		await page.evaluate(() => {
			const root = document.querySelector<HTMLElement>('#native-host')!;
			root.removeAttribute('style');
			window.nativeWord!.handle.element.theme = 'light';
		});
		expect((await snapshot()).canvas).toBe(light.canvas);
		await page.evaluate(() => {
			window.nativeWord!.handle.element.theme = 'auto';
		});
		expect((await snapshot()).canvas).toBe(dark.canvas);
		expect((await snapshot()).paper).toBe('rgb(255, 255, 255)');
		expect(
			await page.evaluate(() => JSON.stringify(window.nativeWord!.handle.element.documentModel)),
		).toBe(model);
		await page.evaluate(() => window.nativeWord!.permission(false));
		await expect(page.getByRole('button', { name: 'Bold', exact: true })).toBeDisabled();
	});
}

test('full Word ribbon reads shared selected, foreground and border roles', async ({ page }) => {
	await openSample(page);
	await page.locator('docx-editor').evaluate((node) => {
		const editor = node as HTMLElementTagNameMap['docx-editor'];
		editor.theme = 'auto';
		editor.parentElement!.setAttribute('data-office-theme', 'dark');
		editor.parentElement!.style.setProperty('--office-selected', 'rgb(101, 102, 103)');
		editor.parentElement!.style.setProperty('--office-foreground', 'rgb(234, 235, 236)');
		editor.parentElement!.style.setProperty('--office-border', 'rgb(141, 142, 143)');
	});
	const bold = page.locator('docx-editor').getByRole('button', { name: 'Italic', exact: true });
	await bold.hover();
	const color = await bold.evaluate((node) => ({
		background: getComputedStyle(node).backgroundColor,
		foreground: getComputedStyle(node).color,
		border: getComputedStyle(node).borderTopColor,
	}));
	expect(color).toEqual({
		background: 'rgb(101, 102, 103)',
		foreground: 'rgb(234, 235, 236)',
		border: 'rgb(141, 142, 143)',
	});
});

test.describe('embedding shadow scopes', () => {
	test.use({ hasTouch: true });
	for (const framework of ['react', 'vue', 'angular', 'svelte', 'solid']) {
		test(`${framework}: nested shadow modes retain inherited fonts and touch geometry`, async ({
			page,
		}) => {
			await page.emulateMedia({ colorScheme: 'light' });
			await page.goto(`/building-blocks.html?framework=${framework}`);
			await expect(page.locator('[data-office-block="toolbar"]')).toBeVisible();
			await page.evaluate(() => {
				const host = document.createElement('div');
				host.id = 'shadow-word-app';
				document.body.append(host);
				const outer = host.attachShadow({ mode: 'open' });
				const scope = document.createElement('section');
				scope.dataset.officeTheme = 'dark';
				scope.style.setProperty('--office-font', 'monospace');
				scope.style.setProperty('--office-target-size', '60px');
				outer.append(scope);
				const innerHost = document.createElement('div');
				scope.append(innerHost);
				const inner = innerHost.attachShadow({ mode: 'open' });
				const innerScope = document.createElement('section');
				innerScope.id = 'nested-word-scope';
				inner.append(innerScope);
				innerScope.append(document.querySelector('#native-host')!);
			});
			const read = () =>
				page.evaluate(() => {
					const editor = window.nativeWord!.handle.element;
					const inner = editor.getRootNode() as ShadowRoot;
					const toolbar = inner.querySelector<HTMLElement>('[data-office-block="toolbar"]')!;
					const button = toolbar.querySelector('button')!;
					const outer = inner.host.getRootNode() as ShadowRoot;
					return {
						canvas: getComputedStyle(editor.shadowRoot!.querySelector('.dve-canvas')!)
							.backgroundColor,
						toolbar: getComputedStyle(toolbar).backgroundColor,
						font: getComputedStyle(button).fontFamily,
						target: getComputedStyle(button).minHeight,
						outerStyles: outer.querySelectorAll('#office-ui-theme').length,
						innerStyles: inner.querySelectorAll('#office-ui-theme').length,
						coarse: matchMedia('(pointer:coarse)').matches,
					};
				});
			await expect.poll(async () => (await read()).canvas).toBe('rgb(31, 41, 55)');
			const dark = await read();
			expect(dark).toMatchObject({
				toolbar: 'rgb(31, 41, 55)',
				font: 'monospace',
				target: '60px',
				outerStyles: 1,
				innerStyles: 1,
				coarse: true,
			});
			await page.evaluate(() =>
				(window.nativeWord!.handle.element.getRootNode() as ShadowRoot)
					.querySelector('#nested-word-scope')!
					.setAttribute('data-office-theme', 'light'),
			);
			const light = await read();
			expect(light.canvas).not.toBe(dark.canvas);
			expect(light.toolbar).not.toBe(dark.toolbar);
			expect(light).toMatchObject({ font: 'monospace', target: '60px' });
			await page.evaluate(() =>
				(window.nativeWord!.handle.element.getRootNode() as ShadowRoot)
					.querySelector('#nested-word-scope')!
					.removeAttribute('data-office-theme'),
			);
			expect((await read()).canvas).toBe(dark.canvas);
			const bold = page.getByRole('button', { name: 'Bold', exact: true });
			await bold.click();
			await expect(bold).toHaveAttribute('aria-pressed', 'true');
			await page.evaluate(() => window.nativeWord!.permission(false));
			await expect(bold).toBeDisabled();
		});
	}
});

import { expect, test } from '@playwright/test';

import { loadDeck } from './support/deck';
test.use({ viewport: { width: 1440, height: 900 } });
test('Office tokens theme actual chrome and shadow controls while legacy inputs stay authoritative', async ({
	page,
}) => {
	await page.emulateMedia({ colorScheme: 'light' });
	await loadDeck(page);
	await page.evaluate(() => {
		// The demos intentionally supply a ViewerTheme. Clear its CSS inputs to exercise the same
		// fallback path hosts receive when they omit that optional prop. Catalog overrides are tested separately.
		for (const element of document.querySelectorAll<HTMLElement>('[style]'))
			for (const key of Array.from(element.style))
				if (key.startsWith('--pptx-') || key.startsWith('--color-'))
					element.style.removeProperty(key);
		const viewer = document.querySelector<HTMLElement>('[data-pptx-theme-root]')!;
		viewer.setAttribute('data-office-theme-probe-root', '');
		const probe = document.createElement('div');
		probe.id = 'office-theme-probe';
		probe.style.cssText =
			'position:fixed;left:12px;bottom:38px;z-index:10000;display:flex;gap:12px';
		const select = document.createElement('pptx-ui-select');
		Reflect.set(select, 'items', [{ value: 'a', label: 'Alpha' }]);
		Reflect.set(select, 'value', 'a');
		const search = document.createElement('pptx-ui-search');
		search.setAttribute('placeholder', 'Theme probe');
		const check = document.createElement('pptx-ui-checkbox');
		check.setAttribute('checked', '');
		probe.append(select, search, check);
		viewer.append(probe);
	});
	const read = () =>
		page.evaluate(() => {
			const viewer = document.querySelector<HTMLElement>('[data-office-theme-probe-root]')!;
			const select = document.querySelector('div#office-theme-probe pptx-ui-select')!;
			const field = select.shadowRoot!.querySelector('button')!;
			const search = document.querySelector('div#office-theme-probe pptx-ui-search')!;
			const check = document.querySelector('div#office-theme-probe pptx-ui-checkbox')!;
			const color = (node: Element) => {
				const style = getComputedStyle(node);
				return {
					background: style.backgroundColor,
					color: style.color,
					border: style.borderTopColor,
					font: style.fontFamily,
				};
			};
			const text = Array.from(
				document.querySelectorAll<HTMLElement>('[data-pptx-viewport] [data-pptx-element] span'),
			)
				.filter((node) => node.textContent?.trim() && !node.querySelector('span'))
				.slice(0, 12)
				.map((node) => ({
					text: node.textContent,
					color: getComputedStyle(node).color,
					font: getComputedStyle(node).fontFamily,
				}));
			return {
				viewer: color(viewer),
				field: color(field),
				search: color(search),
				check: color(check),
				text,
			};
		});
	await expect(page.locator('#office-theme-probe pptx-ui-select')).toBeVisible();
	const initial = await read();
	await page.evaluate(() => (document.documentElement.dataset.officeTheme = 'dark'));
	const dark = await read();
	expect(dark.viewer.background).not.toBe(initial.viewer.background);
	expect(dark.field.background).not.toBe(initial.field.background);
	await page.evaluate(
		() =>
			(document.querySelector<HTMLElement>('[data-office-theme-probe-root]')!.dataset.officeTheme =
				'light'),
	);
	const scopedLight = await read();
	expect(scopedLight.viewer.background).toBe(initial.viewer.background);
	expect(scopedLight.field.background).toBe(initial.field.background);
	await page.evaluate(
		() =>
			(document.querySelector<HTMLElement>('[data-office-theme-probe-root]')!.dataset.officeTheme =
				'dark'),
	);
	const scopedDark = await read();
	expect(scopedDark.viewer.background).toBe(dark.viewer.background);
	await page.evaluate(
		() =>
			(document.querySelector<HTMLElement>('[data-office-theme-probe-root]')!.dataset.officeTheme =
				'auto'),
	);
	await page.emulateMedia({ colorScheme: 'light' });
	expect((await read()).viewer.background).toBe(initial.viewer.background);
	await page.emulateMedia({ colorScheme: 'dark' });
	expect((await read()).viewer.background).toBe(dark.viewer.background);
	await page.evaluate(() => {
		const host = document.querySelector<HTMLElement>('[data-office-theme-probe-root]')!;
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
	for (const [name, part] of Object.entries({
		viewer: custom.viewer,
		field: custom.field,
		search: custom.search,
	})) {
		expect(part.background, name).toBe('rgb(36, 52, 71)');
		expect(part.color, name).toBe('rgb(231, 239, 251)');
		expect(part.font, name).toContain('Courier New');
	}
	expect(custom.field.border).toBe('rgb(97, 121, 146)');
	expect(custom.check.background).toBe('rgb(167, 139, 250)');
	const field = page.locator('#office-theme-probe pptx-ui-select button');
	await page.keyboard.press('Tab');
	await field.focus();
	expect(await field.evaluate((node) => getComputedStyle(node).outlineColor)).toBe(
		'rgb(196, 181, 253)',
	);
	await page.evaluate(() => {
		const host = document.querySelector<HTMLElement>('[data-office-theme-probe-root]')!;
		host.style.setProperty('--pptx-background', '#3e2447');
		host.style.setProperty('--pptx-card', '#3e2447');
		host.style.setProperty('--pptx-foreground', '#f5dffb');
		host.style.setProperty('--pptx-card-foreground', '#f5dffb');
		host.style.setProperty('--pptx-border', '#a877b4');
		host.style.setProperty('--pptx-input', '#a877b4');
		host.style.setProperty('--pptx-primary', '#d8a3ed');
	});
	const legacy = await read();
	for (const part of [legacy.viewer, legacy.field, legacy.search]) {
		expect(part.background).toBe('rgb(62, 36, 71)');
		expect(part.color).toBe('rgb(245, 223, 251)');
	}
	expect(legacy.field.border).toBe('rgb(168, 119, 180)');
	expect(legacy.check.background).toBe('rgb(216, 163, 237)');
	expect(initial.text.length).toBeGreaterThan(0);
	expect(legacy.text).toEqual(initial.text);
	const switchWidth = await page.evaluate(() => {
		const title = document.querySelector<HTMLElement>('pptx-ui-title-bar')!;
		title.style.setProperty('--pptx-switch-width', '50px');
		const toggle = title.shadowRoot!.querySelector('pptx-ui-switch')!;
		return getComputedStyle(toggle).width;
	});
	expect(switchWidth).toBe('50px');
});

import { test, expect, type Page } from '@playwright/test';
import { openDemo } from './demo-page';
import { officeTheme } from './ribbon';

// The viewer's colours come from the shared `--office-*` theme; a host's `--vv-*` custom
// properties override them. The demo workspace sets `--vv-*` for both of its themes, so its
// accent is its own; a host that sets none gets the Office accent and follows the suite theme.

type Paint = { accent: string; surface: string; ink: string };

async function paint(page: Page): Promise<Paint> {
	return page.evaluate(() => {
		const root = document.querySelector('visio-viewer')!.shadowRoot!;
		const color = (el: Element | null | undefined, property: 'color' | 'backgroundColor') =>
			el ? getComputedStyle(el)[property] : 'missing';
		return {
			accent: color(
				root
					.querySelector('office-ui-ribbon')
					?.shadowRoot?.querySelector('[role="tab"][aria-selected="true"]'),
				'color',
			),
			surface: color(root.querySelector('office-ui-status-bar.status'), 'backgroundColor'),
			ink: color(root.querySelector('office-ui-status-item[data-page-status]'), 'color'),
		};
	});
}

async function openThemedDemo(page: Page, theme: 'light' | 'dark'): Promise<void> {
	await page.emulateMedia({ colorScheme: 'light' });
	await page.addInitScript(
		(value) => localStorage.setItem('vitepress-theme-appearance', value),
		theme,
	);
	await openDemo(page);
	await expect(page.locator('visio-viewer svg.paper')).toHaveAttribute(
		'aria-label',
		'Release workflow',
	);
	await expect(page.locator('html')).toHaveAttribute('data-theme', theme);
}

const demoLight: Paint = {
	accent: 'rgb(57, 85, 163)',
	surface: 'rgb(255, 255, 255)',
	ink: 'rgb(31, 31, 31)',
};
const demoDark: Paint = {
	accent: 'rgb(139, 159, 240)',
	surface: 'rgb(27, 29, 32)',
	ink: 'rgb(240, 239, 236)',
};
const officeLight: Paint = {
	accent: 'rgb(37, 99, 235)',
	surface: 'rgb(255, 255, 255)',
	ink: 'rgb(31, 41, 55)',
};
// Accent text is the accent ink: in a dark scheme a lighter tint of the accent, for 4.5:1.
const officeDark: Paint = {
	accent: 'color(srgb 0.632941 0.64 0.967059)',
	surface: 'rgb(17, 24, 39)',
	ink: 'rgb(249, 250, 251)',
};

test('the demo accent and the theme toggle reach the viewer live', async ({ page }) => {
	await openThemedDemo(page, 'light');
	await expect.poll(() => paint(page)).toEqual(demoLight);
	await page
		.locator('visio-viewer')
		.getByRole('tab', { name: 'Architecture', exact: true })
		.click();
	const svg = page.locator('visio-viewer svg.paper');
	await expect(svg).toHaveAttribute('aria-label', 'Architecture');

	// Office Theme in File > Options switches the suite theme; the viewer repaints without reloading.
	await officeTheme(page.locator('visio-viewer'), 'Black');
	await expect(page.locator('html')).toHaveAttribute('data-theme', 'dark');
	await expect.poll(() => paint(page)).toEqual(demoDark);
	await expect(svg).toHaveAttribute('aria-label', 'Architecture');

	await officeTheme(page.locator('visio-viewer'), 'White');
	await expect(page.locator('html')).toHaveAttribute('data-theme', 'light');
	await expect.poll(() => paint(page)).toEqual(demoLight);
});

// `initial` makes a custom property guaranteed-invalid, so every `var(--vv-*, ...)` takes its
// fallback: the page then behaves like a host that sets no `--vv-*` at all.
const unsetHostTokens = `:root, :root[data-theme='light'], :root[data-theme='dark'] {
	--vv-background: initial; --vv-surface: initial; --vv-secondary: initial; --vv-ink: initial;
	--vv-muted: initial; --vv-border: initial; --vv-accent: initial; --vv-accent-ink: initial;
	--vv-accent-soft: initial; --vv-danger: initial; --vv-focus: initial; --vv-shadow: initial;
}`;

for (const theme of ['light', 'dark'] as const) {
	test(`without host tokens the viewer follows the Office theme (${theme} demo)`, async ({
		page,
	}) => {
		await openThemedDemo(page, theme);
		await page.addStyleTag({ content: unsetHostTokens });
		// The demo's data-theme no longer matters; the shared Office theme decides.
		await expect.poll(() => paint(page)).toEqual(officeLight);

		await page.evaluate(() => (document.documentElement.dataset.officeTheme = 'dark'));
		await expect.poll(() => paint(page)).toEqual(officeDark);

		await page.evaluate(() => (document.documentElement.dataset.officeTheme = 'light'));
		await expect.poll(() => paint(page)).toEqual(officeLight);

		// With no explicit choice the Office theme follows the system preference.
		await page.evaluate(() => delete document.documentElement.dataset.officeTheme);
		await page.emulateMedia({ colorScheme: 'dark' });
		await expect.poll(() => paint(page)).toEqual(officeDark);
	});
}

/** The current page tab's label and the shapes search field's boundary. */
async function pageBar(page: Page) {
	return page.evaluate(() => {
		const root = document.querySelector('visio-viewer')!.shadowRoot!;
		const tab = root
			.querySelector('office-ui-tab-strip')!
			.shadowRoot!.querySelector('button.tab[aria-selected="true"]')!;
		const search = root.querySelector('.shapes-search-field');
		return {
			tab: getComputedStyle(tab).color,
			search: search ? getComputedStyle(search).borderBottomColor : 'missing',
		};
	});
}

test('host overrides reach the current page tab in order: --vv-accent-text, --vv-accent, a container accent', async ({
	page,
}) => {
	await openThemedDemo(page, 'light');
	await page.addStyleTag({ content: unsetHostTokens });
	await page.evaluate(() => (document.documentElement.dataset.officeTheme = 'light'));
	await expect.poll(async () => (await pageBar(page)).tab).toBe(officeLight.accent);

	// A container's accent reaches the label: the accent ink is resolved where it is read.
	const container = () =>
		page.evaluate((value) => {
			document
				.querySelector('visio-viewer')!
				.parentElement!.style.setProperty('--office-accent', value);
		}, 'rgb(170, 0, 0)');
	await container();
	await expect.poll(async () => (await pageBar(page)).tab).toBe('rgb(170, 0, 0)');

	const host = (name: string, value: string) =>
		page.evaluate(
			([n, v]) => document.querySelector<HTMLElement>('visio-viewer')!.style.setProperty(n!, v!),
			[name, value],
		);
	await host('--vv-accent', 'rgb(0, 100, 0)');
	await expect.poll(async () => (await pageBar(page)).tab).toBe('rgb(0, 100, 0)');
	await host('--vv-accent-text', 'rgb(0, 0, 160)');
	await expect.poll(async () => (await pageBar(page)).tab).toBe('rgb(0, 0, 160)');

	// Form fields take the control boundary, which a container can override too.
	await page.evaluate(() =>
		document
			.querySelector('visio-viewer')!
			.parentElement!.style.setProperty('--office-control-border', 'rgb(0, 0, 255)'),
	);
	await expect.poll(async () => (await pageBar(page)).search).toBe('rgb(0, 0, 255)');
});

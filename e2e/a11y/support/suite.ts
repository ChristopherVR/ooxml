/**
 * The suite under test: a page with the appearance a spec asks for, loopback-only networking, the
 * focus helpers installed, and a way to open each surface the gates cover.
 */
import { test as base, expect, type Page } from '@playwright/test';
import {
	excelWorkbook,
	powerPointDeck,
	visioDrawing,
	wordDocument,
	type FixtureFile,
} from './fixtures';
import { installA11yHelpers } from './page-helpers';

export type Mode = 'light' | 'dark';
export type Layout = 'office' | 'simple';
export type Product = 'docx' | 'xlsx' | 'pptx' | 'vsdx';
export type Surface = Product | 'files' | 'chat' | 'settings' | 'teams';

export const PRODUCTS: readonly Product[] = ['docx', 'xlsx', 'pptx', 'vsdx'];
export const SURFACES: readonly Surface[] = ['files', 'chat', 'settings', 'teams', ...PRODUCTS];

export const ACCENTS = [
	'blue',
	'purple',
	'teal',
	'green',
	'orange',
	'rose',
	'cranberry',
	'graphite',
] as const;
export type Accent = (typeof ACCENTS)[number];

export const FILES: Record<Product, () => FixtureFile | Promise<FixtureFile>> = {
	docx: wordDocument,
	xlsx: excelWorkbook,
	pptx: powerPointDeck,
	vsdx: visioDrawing,
};

interface Appearance {
	mode: Mode;
	layout: Layout;
	accent: Accent;
}

/** The appearance and start-up state the app reads from storage before it boots. */
function bootState([mode, layout, accent]: [Mode, Layout, Accent]): void {
	(globalThis as { ooxmlSkipStart?: boolean }).ooxmlSkipStart = true;
	try {
		localStorage.setItem('ooxml-suite-theme', JSON.stringify({ mode, accent }));
		const settings = JSON.parse(localStorage.getItem('ooxml-suite-settings') ?? '{}');
		localStorage.setItem('ooxml-suite-settings', JSON.stringify({ ...settings, layout }));
	} catch {
		// Storage is blocked: the app uses its defaults.
	}
}

const LOOPBACK = new Set(['127.0.0.1', 'localhost']);

export const test = base.extend<Appearance>({
	mode: ['light', { option: true }],
	layout: ['office', { option: true }],
	accent: ['blue', { option: true }],
	context: async ({ context, mode, layout, accent }, use) => {
		// The gates measure this build only: nothing leaves the machine.
		await context.route(
			(url) => /^https?:$/u.test(url.protocol) && !LOOPBACK.has(url.hostname),
			(route) => route.abort(),
		);
		await context.addInitScript(installA11yHelpers);
		await context.addInitScript(bootState, [mode, layout, accent] as [Mode, Layout, Accent]);
		await use(context);
	},
});
export { expect };

const READY = /^(Ready|Local|Connected|Saved)/u;

/** Waits for fonts and for every finite transition and animation, so colours are final. */
export async function settle(page: Page): Promise<void> {
	await page.evaluate(async () => {
		await document.fonts.ready;
		// The document's animations include those in shadow trees.
		const finite = document
			.getAnimations()
			.filter((animation) =>
				Number.isFinite(animation.effect?.getComputedTiming().endTime ?? Infinity),
			);
		await Promise.race([
			Promise.all(finite.map((animation) => animation.finished.catch(() => undefined))),
			new Promise((done) => setTimeout(done, 3_000)),
		]);
	});
}

/** Opens the suite's home page and waits until its controls are wired. */
export async function openHome(page: Page, route = '#/'): Promise<void> {
	await page.goto(`/?suite${route}`);
	await page.waitForFunction(() => document.body.dataset.ai !== undefined, null, {
		timeout: 30_000,
	});
}

/** Opens `product`'s fixture in the suite and waits until the editor reports it ready. */
export async function openDocument(page: Page, product: Product): Promise<void> {
	await openHome(page);
	await page.locator('#file-input').setInputFiles(await FILES[product]());
	await page.locator('.document-host:not([hidden])').waitFor({ timeout: 60_000 });
	await page.waitForFunction(
		(ready) => new RegExp(ready, 'u').test(document.getElementById('status')?.textContent ?? ''),
		READY.source,
		{ timeout: 60_000 },
	);
	// The editors finish laying out (fonts, the first paint of the grid or slide) after `Ready`.
	await page.waitForTimeout(800);
	await settle(page);
}

/** Opens one surface of the suite the way a user reaches it. */
export async function openSurface(page: Page, surface: Surface): Promise<void> {
	if (surface === 'files') {
		await openHome(page);
	} else if (surface === 'chat') {
		await openHome(page, '#/chat');
		await page.locator('#chat:not([hidden]) textarea').first().waitFor({ timeout: 30_000 });
	} else if (surface === 'settings') {
		await openHome(page);
		await page.locator('#theme-picker').click();
		await page.locator('office-ui-options-dialog[open]').waitFor({ timeout: 15_000 });
	} else if (surface === 'teams') {
		await openHome(page, '#/teams');
		await page.locator('teams-app textarea').first().waitFor({ timeout: 30_000 });
	} else {
		await openDocument(page, surface);
		return;
	}
	await page.waitForTimeout(500);
	await settle(page);
}

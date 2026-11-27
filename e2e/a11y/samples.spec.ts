/**
 * The samples gallery (`samples/`, built by scripts/build-samples-page.mjs): every sample has a
 * card with a preview, the page passes axe, and each sample opens in its editor from the card's
 * link. An editor may show only the load notices `samples-expected.json` allows that sample: each
 * one is caused by a gap on docs/parity-checklist.md, so a new notice fails here instead of going
 * unseen. Standing notes that only describe what a feature class supports are listed there as
 * informational and are not counted. check-baseline.mjs keeps both lists from growing.
 */
import { readFileSync } from 'node:fs';
import type { Page } from '@playwright/test';
import expected from './samples-expected.json' with { type: 'json' };
import { axeViolations } from './support/axe';
import { expect, settle, test, type Mode } from './support/suite';

type SampleProduct = 'docx' | 'xlsx' | 'pptx' | 'visio';
interface Sample {
	id: string;
	product: SampleProduct;
	title: string;
	file: string;
	size: number;
}
const manifest: { samples: Sample[] } = JSON.parse(
	readFileSync(new URL('../../samples/manifest.json', import.meta.url), 'utf8'),
);
const APP = { docx: 'Word', xlsx: 'Excel', pptx: 'PowerPoint', visio: 'Visio' } as const;
const EDITOR = {
	docx: 'docx-editor',
	xlsx: 'xlsx-editor',
	pptx: '.pptxv',
	visio: 'visio-viewer',
} as const;
const allowed: Record<string, { notice: string; gap: string }[]> = expected.notices;
const informational: readonly string[] = expected.informational;
/** Visio and PowerPoint notices carry a code, so the whole text is compared; Word's are prose. */
const EXACT: Record<SampleProduct, boolean> = { docx: false, xlsx: false, pptx: true, visio: true };
/** The suite's own status line; Visio keeps a `#status` of its own in its shadow tree. */
const STATUS = 'css:light=#status';
const TOAST = '[data-testid="pptx-compat-toast"]';

test.use({ viewport: { width: 1440, height: 900 } });

async function openGallery(page: Page): Promise<void> {
	await page.goto('/samples/');
	await page.locator('.sample-card').first().waitFor();
	await settle(page);
}

/**
 * The compatibility toasts of the open deck. The viewer keeps its handler to itself, so the toasts
 * are the record; it sets them in the same store update as the slides. So this waits for the deck
 * to be on screen (the status bar counts its slides), then until two reads a few frames apart
 * agree, and only then takes the list.
 */
async function deckNotices(page: Page): Promise<string[]> {
	const host = page.locator('.document-host:not([hidden])');
	await expect(host).toContainText(/Slide \d+ of \d+/u, { timeout: 60_000 });
	await expect
		.poll(
			() =>
				page.evaluate(async (selector) => {
					const count = () => document.querySelectorAll(selector).length;
					const before = count();
					for (let frame = 0; frame < 3; frame++)
						await new Promise((done) => requestAnimationFrame(done));
					return count() === before;
				}, TOAST),
			{ message: 'the compatibility toasts settle', timeout: 15_000 },
		)
		.toBe(true);
	return page
		.locator(TOAST)
		.evaluateAll((toasts) =>
			toasts.map((toast) => `${toast.getAttribute('data-code')}: ${toast.textContent?.trim()}`),
		);
}

/**
 * What the open editor reports about the file it loaded, one line per notice. Visio's diagnostics
 * of severity `info` say how the viewer works on any drawing, so they are left out.
 */
async function loadNotices(page: Page, product: SampleProduct): Promise<string[]> {
	if (product === 'pptx') return deckNotices(page);
	return page.evaluate((kind) => {
		const host = document.querySelector('.document-host:not([hidden])');
		if (kind === 'docx') {
			const editor = host?.querySelector('docx-editor') as unknown as {
				documentModel?: { warnings?: string[] };
			} | null;
			return [...(editor?.documentModel?.warnings ?? [])];
		}
		if (kind === 'xlsx') {
			const editor = host?.querySelector('xlsx-editor') as unknown as {
				workbook?: { warnings?: string[] };
			} | null;
			return [...(editor?.workbook?.warnings ?? [])];
		}
		const viewer = host?.querySelector('visio-viewer') as unknown as {
			document?: { diagnostics?: { code: string; severity: string; message: string }[] };
			renderWarnings?: readonly string[];
		} | null;
		return [
			...(viewer?.document?.diagnostics ?? [])
				.filter((item) => item.severity !== 'info')
				.map((item) => `${item.code}: ${item.message}`),
			...(viewer?.renderWarnings ?? []).map((warning) => `render: ${warning}`),
		];
	}, product);
}

test('the gallery shows a card with a preview and both actions for every sample', async ({
	page,
}) => {
	await openGallery(page);
	await expect(page.locator('.sample-card')).toHaveCount(manifest.samples.length);
	for (const sample of manifest.samples) {
		const card = page.locator(`#sample-${sample.id}`);
		await card.scrollIntoViewIfNeeded();
		await expect(card.locator('h3')).toHaveText(sample.title);
		await expect(
			card.getByRole('link', { name: `Open in ${APP[sample.product]}` }),
		).toHaveAttribute('href', `../?sample=${sample.id}`);
		await expect(card.getByRole('link', { name: 'Download' })).toHaveAttribute('href', sample.file);
		const preview = card.locator('.sample-shot img');
		await expect
			.poll(() => preview.evaluate((img: HTMLImageElement) => img.complete && img.naturalWidth))
			.toBe(800);
	}
	// The files are fetched on demand only: showing the gallery requests none of them.
	const files = await page.evaluate(() =>
		performance
			.getEntriesByType('resource')
			.map((entry) => entry.name)
			.filter((name) => /\.(docx|xlsx|pptx|vsdx)$/u.test(name)),
	);
	expect(files).toEqual([]);
});

test('the page title is the first heading, and the headings go down one level at a time', async ({
	page,
}) => {
	await openGallery(page);
	const levels = await page
		.locator('h1, h2, h3, h4, h5, h6, [role="heading"]')
		.evaluateAll((headings) => headings.map((heading) => Number(heading.tagName.slice(1))));
	expect(levels[0]).toBe(1);
	expect(levels.filter((level) => level === 1)).toHaveLength(1);
	for (const [at, level] of levels.entries())
		expect(level - (levels[at - 1] ?? 0), `heading ${at + 1}`).toBeLessThanOrEqual(1);
});

test('a download link serves the file the manifest describes', async ({ page }) => {
	await openGallery(page);
	const sample = manifest.samples[0]!;
	const [download] = await Promise.all([
		page.waitForEvent('download'),
		page.locator(`#sample-${sample.id}`).getByRole('link', { name: 'Download' }).click(),
	]);
	expect(download.suggestedFilename()).toBe(sample.file.split('/')[1]);
	expect(readFileSync(await download.path()).length).toBe(sample.size);
});

for (const mode of ['light', 'dark'] satisfies Mode[]) {
	test.describe(`gallery, ${mode}`, () => {
		test.use({ mode });
		test('has no critical or serious axe violation', async ({ page }) => {
			await openGallery(page);
			// Bring every card in, so the reveal has finished and lazy previews are in the tree.
			await page.evaluate(() => window.scrollTo(0, document.body.scrollHeight));
			await page.waitForTimeout(900);
			await settle(page);
			expect(await axeViolations(page)).toEqual([]);
		});
	});
}

test.describe('gallery on a phone', () => {
	test.use({
		viewport: { width: 390, height: 844 },
		hasTouch: true,
		isMobile: true,
		deviceScaleFactor: 2,
	});
	test('the open and download targets are at least 44 px high', async ({ page }) => {
		await openGallery(page);
		const heights = await page
			.locator('.sample-card .hub-links a')
			.evaluateAll((links) => links.map((link) => link.getBoundingClientRect().height));
		expect(heights).toHaveLength(manifest.samples.length * 2);
		expect(Math.min(...heights)).toBeGreaterThanOrEqual(44);
	});
});

for (const sample of manifest.samples) {
	test(`${sample.file} opens in ${APP[sample.product]} from its card`, async ({ page }) => {
		test.slow();
		const errors: string[] = [];
		page.on('pageerror', (error) => errors.push(error.message));
		await openGallery(page);
		await page
			.locator(`#sample-${sample.id}`)
			.getByRole('link', { name: `Open in ${APP[sample.product]}` })
			.click();
		const name = sample.file.split('/')[1]!;
		const host = page.locator('.document-host:not([hidden])');
		await host.waitFor({ timeout: 90_000 });
		await expect(page.locator(STATUS)).toHaveText(/^Ready/u, { timeout: 90_000 });
		await expect(host).toHaveAttribute('aria-label', name);
		await expect(host.locator(EDITOR[sample.product]).first()).toBeVisible();
		await expect(page.locator('.file-tab .tab-name')).toHaveText(name);
		// The parameter is taken out of the address, so a reload does not add another copy.
		expect(new URL(page.url()).searchParams.has('sample')).toBe(false);
		// A Word sample is a designed page, so it opens as pages and the Pages pane has thumbnails.
		if (sample.product === 'docx') {
			await expect(host.locator('[data-print-view]')).toBeVisible({ timeout: 30_000 });
			await expect(host).not.toContainText('Page thumbnails need Print Layout');
		}
		const found = (await loadNotices(page, sample.product)).filter(
			(notice) => !informational.some((start) => notice.startsWith(start)),
		);
		const mine = (allowed[sample.id] ?? []).map((entry) => entry.notice);
		const matches = (notice: string, entry: string) =>
			EXACT[sample.product] ? notice === entry : notice.startsWith(entry);
		const unexpected = found.filter((notice) => !mine.some((entry) => matches(notice, entry)));
		expect(unexpected, `load notices of ${sample.file} that are not expected`).toEqual([]);
		const closed = mine.filter((entry) => !found.some((notice) => matches(notice, entry)));
		if (closed.length)
			test.info().annotations.push({
				type: 'expected notices can shrink',
				description: `${sample.id}: ${closed.join(' | ')}`,
			});
		expect(errors, 'page errors').toEqual([]);
	});
}

test('opening a sample twice reuses the untouched copy', async ({ page }) => {
	const sample = manifest.samples.find((item) => item.product === 'visio') ?? manifest.samples[0]!;
	const name = sample.file.split('/')[1]!;
	for (let visit = 0; visit < 2; visit++) {
		await page.goto(`/?sample=${sample.id}`);
		await page.locator('.document-host:not([hidden])').waitFor({ timeout: 90_000 });
		await expect(page.locator(STATUS)).toHaveText(/^Ready/u, { timeout: 90_000 });
	}
	await page.goto('/?suite#/');
	await expect(page.locator('#files [data-open]').filter({ hasText: name })).toHaveCount(1);
});

for (const [what, id] of [
	['an unknown sample', 'no-such-sample'],
	['an id that cannot be one', '../index'],
] as const) {
	test(`${what} opens nothing and says so beside a link back to the gallery`, async ({ page }) => {
		await page.goto(`/?sample=${encodeURIComponent(id)}`);
		const notice = page.locator('#sample-notice');
		await expect(notice).toBeVisible();
		await expect(notice).toContainText('That sample is not available.');
		// The same words reach assistive technology through the status line, once.
		await expect(page.locator(STATUS)).toHaveText('That sample is not available.');
		await expect(page.locator('.document-host:not([hidden])')).toHaveCount(0);
		// The workspace is on its usual route, as if the address had named no sample.
		await expect(page.locator('#home:not([hidden]), #chat:not([hidden])')).toBeVisible();
		expect(new URL(page.url()).searchParams.has('sample')).toBe(false);
		await notice.getByRole('link', { name: 'Browse samples' }).click();
		await expect(page).toHaveURL(/\/samples\/$/u);
		await expect(page.locator('.sample-card').first()).toBeVisible();
	});
}

test('the start screen offers the sample files', async ({ browser, baseURL }) => {
	// A context of its own: the shared one skips the start screen for the workspace specs.
	const context = await browser.newContext({ ...(baseURL ? { baseURL } : {}) });
	try {
		const page = await context.newPage();
		await page.goto('/');
		await page.locator('#start-screen [data-start="samples"]').click();
		await expect(page).toHaveURL(/\/samples\/$/u);
		await expect(page.locator('h1')).toHaveText('Sample files');
	} finally {
		await context.close();
	}
});

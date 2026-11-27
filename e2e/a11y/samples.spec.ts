/**
 * The samples gallery (`samples/`, built by scripts/build-samples-page.mjs): every sample has a
 * card with a preview, the page passes axe, and each sample opens in its editor from the card's
 * link. An editor may show only the load notices listed in `samples-expected.json`: each one is a
 * parity gap the sample shows on purpose (docs/parity-checklist.md), so a new notice fails here
 * instead of going unseen, and the list shrinks as gaps close. An entry there is the start of a
 * notice (its code for PowerPoint and Visio, its opening words for Word and Excel), so rewording
 * the rest of a notice does not fail the gate.
 */
import { readFileSync } from 'node:fs';
import type { Page } from '@playwright/test';
import expected from './samples-expected.json' with { type: 'json' };
import { axeViolations } from './support/axe';
import { expect, settle, test, type Mode } from './support/suite';

interface Sample {
	id: string;
	product: 'docx' | 'xlsx' | 'pptx' | 'visio';
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
const notices: Record<string, string[]> = expected.notices;
/** The suite's own status line; Visio keeps a `#status` of its own in its shadow tree. */
const STATUS = 'css:light=#status';

test.use({ viewport: { width: 1440, height: 900 } });

async function openGallery(page: Page): Promise<void> {
	await page.goto('/samples/');
	await page.locator('.sample-card').first().waitFor();
	await settle(page);
}

/** What the open editor reports about the file it loaded, one line per notice. */
async function loadNotices(page: Page, product: Sample['product']): Promise<string[]> {
	if (product === 'pptx') {
		// The viewer keeps its handler to itself; its compatibility toasts carry the codes.
		await page.waitForTimeout(1_000);
		return page
			.locator('[data-testid="pptx-compat-toast"]')
			.evaluateAll((toasts) =>
				toasts.map((toast) => `${toast.getAttribute('data-code')}: ${toast.textContent?.trim()}`),
			);
	}
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
			document?: { diagnostics?: { code: string; message: string }[] };
			renderWarnings?: readonly string[];
		} | null;
		return [
			...(viewer?.document?.diagnostics ?? []).map((item) => `${item.code}: ${item.message}`),
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
		const found = await loadNotices(page, sample.product);
		const allowed = notices[sample.id] ?? [];
		const unexpected = found.filter((notice) => !allowed.some((start) => notice.startsWith(start)));
		expect(unexpected, `load notices of ${sample.file} that are not expected`).toEqual([]);
		const closed = allowed.filter((start) => !found.some((notice) => notice.startsWith(start)));
		if (closed.length)
			test.info().annotations.push({
				type: 'expected notices can shrink',
				description: `${sample.id}: ${closed.join(' | ')}`,
			});
		expect(errors, 'page errors').toEqual([]);
	});
}

test('opening a sample twice reuses the untouched copy, and an unknown id opens nothing', async ({
	page,
}) => {
	const sample = manifest.samples.find((item) => item.product === 'visio') ?? manifest.samples[0]!;
	const name = sample.file.split('/')[1]!;
	for (let visit = 0; visit < 2; visit++) {
		await page.goto(`/?sample=${sample.id}`);
		await page.locator('.document-host:not([hidden])').waitFor({ timeout: 90_000 });
		await expect(page.locator(STATUS)).toHaveText(/^Ready/u, { timeout: 90_000 });
	}
	await page.goto('/?suite#/');
	await expect(page.locator('#files [data-open]').filter({ hasText: name })).toHaveCount(1);
	await page.goto('/?sample=no-such-sample');
	await expect(page.locator(STATUS)).toHaveText('That sample is not available.');
	await expect(page.locator('.document-host:not([hidden])')).toHaveCount(0);
});

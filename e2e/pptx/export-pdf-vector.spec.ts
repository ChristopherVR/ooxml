/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright spec, `test`/`expect` come from @playwright/test */
/**
 * File > Export > Create PDF writes vector pages by default, like PowerPoint's Save as PDF:
 * one page per visible slide at the slide's own size, text as real PDF text operators (so it is
 * selectable and searchable), hidden slides left out unless Options > Advanced > Save as PDF >
 * "Include hidden slides" is on, and the picture export still available as the compatibility
 * option. The file's streams are inflated and read here, so a binding that silently fell back
 * to one picture per page fails.
 *
 * Run: bunx playwright test export-pdf-vector
 */
import { readFile } from 'node:fs/promises';
import { inflateSync } from 'node:zlib';

import { expect, test } from '@playwright/test';
import type { Page } from '@playwright/test';
import { PptxHandler } from 'pptx-viewer-core';

import { fixture, loadDeck, loadDeckAt } from './support/deck';
import {
	downloadBytes,
	downloadViaCard,
	EXPORT_DECK,
	isPdf,
	openBackstageExport,
	PDF_CARD,
	pdfImageXObjectCount,
	pdfOptionsDialog,
	pdfPageCount,
	seedAdvancedOptions,
	type PdfOptionsChoice,
} from './support/exports';
import { byBinding } from './support/menu-report';
import { acrossFrameworks } from './support/parity';

const VIEWPORT = { width: 1600, height: 950 };
test.use({ viewport: VIEWPORT });
test.describe.configure({ timeout: 240_000 });

/** A deck with one hidden slide among fourteen. */
const HIDDEN_DECK = fixture('solution-explorer.pptx');
/** A deck with picture effects, so the export approximates something. */
const NOTICE_DECK = fixture('image-alpha-effects.pptx');

interface DeckFacts {
	slides: number;
	visible: number;
	/** Page size in points from the slide size. */
	width: number;
	height: number;
}

async function deckFacts(path: string): Promise<DeckFacts> {
	const data = await new PptxHandler().load(new Uint8Array(await readFile(path)).buffer);
	return {
		slides: data.slides.length,
		visible: data.slides.filter((slide) => !slide.hidden).length,
		width: (data.width * 72) / 96,
		height: (data.height * 72) / 96,
	};
}

/** Every content and font stream of the file, inflated. */
function pdfStreams(bytes: Uint8Array): string[] {
	const text = Buffer.from(bytes).toString('latin1');
	const out: string[] = [];
	for (const match of text.matchAll(/<<([^]*?)>>\nstream\n/gu)) {
		const length = Number(/\/Length (\d+)/u.exec(match[1]!)?.[1]);
		const start = match.index! + match[0].length;
		const data = Buffer.from(bytes.subarray(start, start + length));
		if (!/\/FlateDecode/u.test(match[1]!) || /\/Subtype \/Image/u.test(match[1]!)) continue;
		try {
			out.push(inflateSync(data).toString('latin1'));
		} catch {
			// A stream this reader cannot inflate is not a content stream.
		}
	}
	return out;
}

interface PdfFacts {
	isPdf: boolean;
	pages: number;
	images: number;
	textRuns: number;
	mediaBox: string | undefined;
	text: string;
}

function readPdf(bytes: Uint8Array): PdfFacts {
	const content = pdfStreams(bytes).join('\n');
	const literal = [...content.matchAll(/\((.*?)\) Tj/gu)].map((m) => m[1]).join(' ');
	return {
		isPdf: isPdf(bytes),
		pages: pdfPageCount(bytes),
		images: pdfImageXObjectCount(bytes),
		textRuns: (content.match(/\bT[jJ]\b/gu) ?? []).length,
		mediaBox: /\/MediaBox \[([^\]]*)\]/u.exec(Buffer.from(bytes).toString('latin1'))?.[1],
		text: literal,
	};
}

async function exportPdf(page: Page, choice: PdfOptionsChoice = {}): Promise<PdfFacts> {
	await openBackstageExport(page);
	const download = await downloadViaCard(page, PDF_CARD, 180_000, choice);
	expect(download.suggestedFilename()).toMatch(/\.pdf$/u);
	return readPdf(await downloadBytes(download));
}

test.describe('vector PDF export', () => {
	test('writes slide-sized pages with real text, one per slide', async ({ page }) => {
		const facts = await deckFacts(EXPORT_DECK);
		await loadDeck(page, EXPORT_DECK);
		const pdf = await exportPdf(page);
		expect(pdf.isPdf).toBe(true);
		expect(pdf.pages).toBe(facts.visible);
		expect(pdf.mediaBox).toBe(`0 0 ${facts.width} ${facts.height}`);
		expect(pdf.textRuns, 'slide text must be PDF text operators').toBeGreaterThan(facts.visible);
		expect(pdf.images, 'pages must not be one picture each').toBeLessThan(facts.visible);
	});

	test('leaves hidden slides out unless the option includes them', async ({ page }) => {
		const facts = await deckFacts(HIDDEN_DECK);
		expect(facts.visible).toBeLessThan(facts.slides);
		await loadDeck(page, HIDDEN_DECK);
		expect((await exportPdf(page)).pages).toBe(facts.visible);
	});

	test('includes hidden slides when Options > Advanced asks for them', async ({ page }) => {
		const facts = await deckFacts(HIDDEN_DECK);
		await seedAdvancedOptions(page, { pdfIncludeHiddenSlides: true });
		await loadDeck(page, HIDDEN_DECK);
		expect((await exportPdf(page)).pages).toBe(facts.slides);
	});

	test('the options step exports the current slide or a range', async ({ page }) => {
		const facts = await deckFacts(HIDDEN_DECK);
		await loadDeck(page, HIDDEN_DECK);
		expect((await exportPdf(page, { range: 'current' })).pages).toBe(1);
		// Slides 1-3 counted with hidden slides in the numbering, all included.
		expect(
			(await exportPdf(page, { range: 'slides', from: 1, to: 3, includeHidden: true })).pages,
		).toBe(Math.min(3, facts.slides));
	});

	test('the options step can be cancelled without a download', async ({ page }) => {
		await loadDeck(page, EXPORT_DECK);
		await openBackstageExport(page);
		const downloads: string[] = [];
		page.on('download', (download) => downloads.push(download.suggestedFilename()));
		await page
			.locator('[role="dialog"][aria-label="File"]')
			.getByRole('button', { name: PDF_CARD })
			.first()
			.click();
		await expect(pdfOptionsDialog(page)).toBeVisible();
		await page.keyboard.press('Escape');
		await expect(page.locator('pptx-ui-pdf-export-dialog')).toHaveCount(0);
		await page.waitForTimeout(1_000);
		expect(downloads).toEqual([]);
	});

	test(
		'every binding shows one notice exactly when the PDF approximated something',
		{ tag: '@cross-binding' },
		async ({ browser }, testInfo) => {
			const results = await acrossFrameworks(
				browser,
				testInfo,
				async (page, origin) => {
					let approximated = false;
					page.on('console', (message) => {
						if (message.text().startsWith('[pptx-viewer] PDF export:')) approximated = true;
					});
					await loadDeckAt(page, origin, NOTICE_DECK);
					await exportPdf(page);
					const notice = page.locator(
						'[data-testid="pptx-compat-toast"][data-code="PDF_EXPORT_APPROXIMATED"]',
					);
					if (approximated)
						await notice
							.first()
							.waitFor({ timeout: 10_000 })
							.catch(() => {});
					return {
						approximated,
						notices: await notice.count(),
						details: await notice.locator('details li').count(),
						message:
							(await notice
								.first()
								.locator('.message')
								.textContent()
								.catch(() => '')) ?? '',
					};
				},
				{ viewport: VIEWPORT },
			);
			const problems = byBinding(results).flatMap(({ name, value }) => {
				if (!value.approximated)
					return value.notices ? [`${name}: a notice without approximations`] : [];
				const issues: string[] = [];
				if (value.notices !== 1) issues.push(`${name}: ${value.notices} PDF notices, expected 1`);
				if (!/^PDF saved with approximations: .+\.$/u.test(value.message))
					issues.push(`${name}: notice text "${value.message}"`);
				if (value.details < 1) issues.push(`${name}: the notice lists no details`);
				return issues;
			});
			expect(problems.join('\n')).toBe('');
		},
	);

	test('keeps the picture export as the compatibility option', async ({ page }) => {
		const facts = await deckFacts(EXPORT_DECK);
		await seedAdvancedOptions(page, { pdfSlidesAsPictures: true });
		await loadDeck(page, EXPORT_DECK);
		const pdf = await exportPdf(page);
		expect(pdf.pages).toBe(facts.slides);
		expect(pdf.images).toBeGreaterThanOrEqual(facts.slides);
	});

	test(
		'every binding reads Options > Advanced > Save as PDF',
		{ tag: '@cross-binding' },
		async ({ browser }, testInfo) => {
			const facts = await deckFacts(EXPORT_DECK);
			const results = await acrossFrameworks(
				browser,
				testInfo,
				async (page, origin) => {
					await seedAdvancedOptions(page, { pdfSlidesAsPictures: true });
					await loadDeckAt(page, origin, EXPORT_DECK);
					return exportPdf(page);
				},
				{ viewport: VIEWPORT },
			);
			const problems = byBinding(results).flatMap(({ name, value }) =>
				value.pages === facts.slides && value.images >= facts.slides
					? []
					: [
							`${name}: "Save slides as pictures" gave ${value.pages} pages, ${value.images} images`,
						],
			);
			expect(problems.join('\n')).toBe('');
		},
	);

	test(
		'every binding writes vector pages with text',
		{ tag: '@cross-binding' },
		async ({ browser }, testInfo) => {
			const facts = await deckFacts(EXPORT_DECK);
			const results = await acrossFrameworks(
				browser,
				testInfo,
				async (page, origin) => {
					await loadDeckAt(page, origin, EXPORT_DECK);
					return exportPdf(page);
				},
				{ viewport: VIEWPORT },
			);
			const problems = byBinding(results).flatMap(({ name, value }) => {
				const issues: string[] = [];
				if (value.pages !== facts.visible)
					issues.push(`${name}: ${value.pages} pages, expected ${facts.visible}`);
				if (value.textRuns <= facts.visible) issues.push(`${name}: no PDF text operators`);
				if (value.images >= facts.visible) issues.push(`${name}: pages are pictures`);
				return issues;
			});
			expect(problems.join('\n')).toBe('');
		},
	);
});

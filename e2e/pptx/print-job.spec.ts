/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright spec */
/**
 * File > Print in all five bindings: the shared print job plan.
 *
 * Every binding used to decide on its own which slides a confirmed Print prints, how it renders
 * them and what the print window is called, and they drifted: React never took its vector path
 * (it captured every slide of the deck as an image), Vue and vanilla titled the window "Slides
 * (Vector)" in every language. With the defaults (all slides, full-page slides, hidden slides
 * left out) every binding now prints one vector page per visible slide under the "Slides" title.
 *
 * The print window is read as each binding opens it: a popup in four bindings, a hidden iframe
 * (`srcdoc`) in Svelte. `window.print` is stubbed, so nothing reaches a printer dialog.
 *
 * Run: bunx playwright test print-job
 */
import { readFile } from 'node:fs/promises';

import { expect, test } from '@playwright/test';
import type { Page } from '@playwright/test';
import { PptxHandler } from 'pptx-viewer-core';

import { loadDeck, ribbon, SAMPLE_DECK } from './support/deck';

test.use({ viewport: { width: 1440, height: 900 } });

/** How many slides of a deck may print by default: hidden ones are left out, as in PowerPoint. */
async function visibleSlideCount(path: string): Promise<number> {
	const data = await new PptxHandler().load(new Uint8Array(await readFile(path)).buffer);
	return data.slides.filter((slide) => !slide.hidden).length;
}

/** File > Print > Print Presentation: the binding's Print dialog. */
async function openPrintDialog(page: Page) {
	await ribbon(page).getByRole('tab', { name: 'File', exact: true }).click();
	const shell = page.locator('[role="dialog"][aria-label="File"]');
	await shell.waitFor();
	await shell.locator('[data-pptx-backstage-nav-item]', { hasText: 'Print' }).first().click();
	await page
		.getByRole('button', { name: /Print Presentation/u })
		.first()
		.click();
	const dialog = page.locator('[role="dialog"]').filter({ hasText: /print/iu }).last();
	await expect(dialog).toBeVisible();
	return dialog;
}

/** Confirm Print and return the document it writes, from the popup or the hidden iframe. */
async function confirmAndReadPrintDocument(page: Page): Promise<string> {
	const dialog = await openPrintDialog(page);
	// Whichever the binding uses arrives first; the other one times out quietly.
	const popup = page.waitForEvent('popup', { timeout: 30_000 }).catch(() => null);
	const frame = page.locator('iframe[srcdoc]').last();
	const framed = frame
		.waitFor({ state: 'attached', timeout: 30_000 })
		.then(() => 'iframe' as const)
		.catch(() => null);
	await dialog
		.getByRole('button', { name: /^print$/iu })
		.last()
		.click();
	const opened = await Promise.race([popup, framed]);
	if (opened === 'iframe') {
		return (await frame.getAttribute('srcdoc')) ?? '';
	}
	expect(opened, 'Print must open a print window or a print frame').not.toBeNull();
	if (!opened) {
		return '';
	}
	let html = '';
	await expect
		.poll(
			async () => {
				html = await opened.content();
				return /<\/body>/u.test(html) && /class="(?:print-slide-page|page slide-page)"/u.test(html);
			},
			{ message: 'the print window must receive its document', timeout: 30_000 },
		)
		.toBeTruthy();
	return html;
}

test.describe('print job', () => {
	test('prints every visible slide as one vector page under the Slides title', async ({ page }) => {
		await page.addInitScript(() => {
			window.print = () => {};
		});
		const expected = await visibleSlideCount(SAMPLE_DECK);
		await loadDeck(page, SAMPLE_DECK);

		const html = await confirmAndReadPrintDocument(page);

		expect(/<title>([^<]*)<\/title>/u.exec(html)?.[1], 'the print window title').toBe('Slides');
		expect(
			html.match(/class="print-slide-page"/gu) ?? [],
			'one vector page per visible slide',
		).toHaveLength(expected);
		expect(html, 'no slide printed as a captured image').not.toMatch(/class="slide-img"/u);
	});
});

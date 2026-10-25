/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright spec */
/**
 * A shape's hyperlink saves the target the editor holds, through each binding's own state and
 * the File backstage save: an address edited after load, a slide jump made in the editor, and a
 * linked shape copied onto another slide (its loaded relationship id names nothing there).
 */
import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';

import { expect, test } from '@playwright/test';
import type { Page } from '@playwright/test';
import JSZip from 'jszip';
import type { PowerPointViewerAPI } from 'pptx-viewer-shared';

import { savePptxViaBackstage } from './save-pptx';
import { resetTabSession, slideStage } from './support/deck';

declare global {
	interface Window {
		/** Public component handle exposed by each development demo. */
		__pptxViewer: PowerPointViewerAPI;
	}
}

// One slide; its rectangle links to https://example.com/hyperlink-fixture.
const fixture = fileURLToPath(new URL('./fixtures/hyperlink-action.pptx', import.meta.url));
const LOADED_URL = 'https://example.com/hyperlink-fixture';

async function open(page: Page): Promise<void> {
	await resetTabSession(page);
	await page.goto('/');
	await page.locator('#file-input').setInputFiles(fixture);
	await expect(slideStage(page)).toBeVisible();
	await expect
		.poll(() =>
			page.evaluate(() => window.__pptxViewer.getElements(0).some((el) => el.actionClick)),
		)
		.toBe(true);
}

/** A blank second slide, with the first one active again. */
async function addSecondSlide(page: Page): Promise<void> {
	await page.evaluate(() => window.__pptxViewer.addSlide(0));
	await expect.poll(() => page.evaluate(() => window.__pptxViewer.getSlideCount())).toBe(2);
}

async function activate(page: Page, index: number): Promise<void> {
	await page.evaluate((i) => window.__pptxViewer.setActiveSlideIndex(i), index);
	await expect
		.poll(() => page.evaluate(() => window.__pptxViewer.getActiveSlideIndex()))
		.toBe(index);
}

/** Save through File, then the target of slide `number`'s `a:hlinkClick`, as its relationship names it. */
async function savedTarget(page: Page, number: number): Promise<string | undefined> {
	const download = await savePptxViaBackstage(page);
	const zip = await JSZip.loadAsync(await readFile((await download.path())!));
	const xml = await zip.file(`ppt/slides/slide${number}.xml`)!.async('string');
	const rels = await zip.file(`ppt/slides/_rels/slide${number}.xml.rels`)!.async('string');
	const rId = /<a:hlinkClick[^>]*r:id="([^"]*)"/u.exec(xml)?.[1];
	const relationship = rId
		? new RegExp(`<Relationship [^>]*Id="${rId}"[^>]*>`, 'u').exec(rels)?.[0]
		: undefined;
	return relationship ? /Target="([^"]*)"/u.exec(relationship)?.[1] : undefined;
}

test('an address edited after load saves the new address', async ({ page }) => {
	await open(page);
	await page.evaluate(() => {
		const api = window.__pptxViewer;
		const el = api.getElements(0).find((element) => element.actionClick)!;
		// What the hyperlink dialog writes: the loaded action, with a new address.
		api.updateElement(el.id, {
			actionClick: { ...el.actionClick, url: 'https://example.org/changed' },
		});
	});
	await expect
		.poll(() =>
			page.evaluate(
				() => window.__pptxViewer.getElements(0).find((el) => el.actionClick)?.actionClick?.url,
			),
		)
		.toBe('https://example.org/changed');
	expect(await savedTarget(page, 1)).toBe('https://example.org/changed');
});

test('a slide jump made in the editor saves the slide it points at', async ({ page }) => {
	await open(page);
	await addSecondSlide(page);
	await activate(page, 0);
	await page.evaluate(() => {
		const api = window.__pptxViewer;
		const el = api.getElements(0).find((element) => element.actionClick)!;
		// What every Action Settings panel writes for "Hyperlink to: Slide 2".
		api.updateElement(el.id, {
			actionClick: { action: 'ppaction://hlinksldjump', targetSlideIndex: 1 },
		});
	});
	await expect
		.poll(() =>
			page.evaluate(
				() =>
					window.__pptxViewer.getElements(0).find((el) => el.actionClick)?.actionClick
						?.targetSlideIndex,
			),
		)
		.toBe(1);
	expect(await savedTarget(page, 1)).toBe('slide2.xml');
});

test('a linked shape copied onto another slide saves its address there', async ({ page }) => {
	await open(page);
	await addSecondSlide(page);
	await activate(page, 1);
	await page.evaluate(() => {
		const api = window.__pptxViewer;
		api.addElement(api.getElements(0).find((element) => element.actionClick)!);
	});
	await expect
		.poll(() =>
			page.evaluate(() => window.__pptxViewer.getElements(1).some((el) => el.actionClick)),
		)
		.toBe(true);
	expect(await savedTarget(page, 2)).toBe(LOADED_URL);
});

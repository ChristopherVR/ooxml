/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright spec */
/**
 * OneDrive under File > Open in every binding, against a test-only Graph stand-in: sign in from the
 * OneDrive place, open a deck from the listing, edit, and the viewer's own File > Save uploads to the
 * same item. The status bar reads "All saved" only once the upload has landed; a failed upload leaves
 * it unsaved; a file on another drive opens read-only.
 */
import { readFile } from 'node:fs/promises';

import { expect, test } from '@playwright/test';
import type { Page } from '@playwright/test';
import type { PowerPointViewerAPI } from 'pptx-viewer-shared';

import { loadDeck, openRibbonTab, slideElements, thumbnail } from './support/deck';
import { pressShortcut } from './support/keyboard';
import { ONEDRIVE_TEST_TOKEN, routeOneDriveGraph } from './support/onedrive-graph';

declare global {
	interface Window {
		/** Public component handle exposed by each development demo. */
		__pptxViewer: PowerPointViewerAPI;
	}
}

const deck = () => readFile(new URL('./fixtures/sample-deck.pptx', import.meta.url));
const unsaved = (page: Page) => page.getByText('Unsaved changes', { exact: true });

async function openBackstage(page: Page) {
	const toolbar = page.getByRole('toolbar', { name: 'Presentation toolbar' });
	await toolbar.getByRole('tab', { name: 'File', exact: true }).click();
	return page.getByRole('dialog', { name: 'File' });
}

/**
 * File > Open > OneDrive (signing in first if the place asks) > open `name`, on a viewer that
 * already has a document.
 */
async function openInBackstage(page: Page, name: string): Promise<void> {
	const backstage = await openBackstage(page);
	// The nav entry, not the cloud browser's own Open button (a place can stay selected).
	await backstage.locator('[data-pptx-backstage-nav-item]', { hasText: /^\s*Open\s*$/u }).click();
	await backstage.getByRole('tab', { name: 'OneDrive', exact: true }).click();
	const signIn = backstage.getByRole('button', { name: 'Sign in', exact: true });
	const options = backstage.getByRole('option').first();
	await expect(signIn.or(options)).toBeVisible();
	if (await signIn.isVisible()) {
		await signIn.click();
		const dialog = page.getByRole('dialog', { name: 'Connect OneDrive (demo)' });
		await dialog.getByLabel('Access token').fill(ONEDRIVE_TEST_TOKEN);
		await dialog.getByRole('button', { name: 'Continue', exact: true }).click();
	}
	await backstage.getByRole('option', { name: new RegExp(name, 'iu') }).click();
	// Opening closes File and replaces the deck that was showing (the title bar names the file).
	await expect(backstage).toBeHidden({ timeout: 30_000 });
	await expect(page.getByText(new RegExp(`${name}\\.pptx`, 'iu')).first()).toBeVisible();
	await expect(page.locator('[data-pptx-viewport] [data-pptx-element="true"]').first()).toBeVisible(
		{ timeout: 30_000 },
	);
}

/** Load a device deck, then File > Open > OneDrive > `name`. */
async function openFromOneDrive(page: Page, name: string): Promise<void> {
	// The viewer (and its File tab) needs a document first.
	await loadDeck(page);
	await openInBackstage(page, name);
}

interface HandleEdits {
	/** The deck as it was before any attempt. */
	slides: number;
	elements: number;
	firstX: number;
	/** `addElement`'s result: the new id, or undefined when it refused. */
	inserted: string | undefined;
	insertedElements: number;
	/** What each attempt changed on the active slide (and the slide count for `addSlide`). */
	updated: boolean;
	duplicated: number;
	deleted: number;
	slidesAdded: number;
}

/**
 * Every edit the public handle offers, tried in turn on the active slide, and what each did. Each
 * call gets a paint before it is read back: some bindings publish a handle that reads the last
 * render, so a read in the same tick as the write sees the old deck.
 */
async function tryHandleEdits(page: Page): Promise<HandleEdits> {
	const settle = () => page.waitForTimeout(400);
	const read = () =>
		page.evaluate(() => {
			const api = window.__pptxViewer;
			return {
				slides: api.getSlideCount(),
				elements: api.getElements(),
				active: api.getActiveSlideIndex(),
			};
		});
	const before = await read();
	const first = before.elements[0]!;
	const inserted = await page.evaluate(() =>
		window.__pptxViewer.addElement({
			type: 'shape',
			id: 'probe',
			x: 10,
			y: 10,
			width: 80,
			height: 80,
		} as never),
	);
	await settle();
	const afterInsert = await read();
	await page.evaluate((id) => window.__pptxViewer.updateElement(id, { x: 555 } as never), first.id);
	await settle();
	const afterUpdate = await read();
	await page.evaluate((id) => window.__pptxViewer.duplicateElement(id), first.id);
	await settle();
	const afterDuplicate = await read();
	await page.evaluate((id) => window.__pptxViewer.deleteElements([id]), first.id);
	await settle();
	const afterDelete = await read();
	await page.evaluate(() => window.__pptxViewer.addSlide(0));
	await settle();
	const afterAddSlide = await read();
	return {
		slides: before.slides,
		elements: before.elements.length,
		firstX: first.x,
		inserted,
		insertedElements: afterInsert.elements.length - before.elements.length,
		updated: afterUpdate.elements.find((e) => e.id === first.id)?.x !== first.x,
		duplicated: afterDuplicate.elements.length - afterUpdate.elements.length,
		deleted: afterDuplicate.elements.length - afterDelete.elements.length,
		slidesAdded: afterAddSlide.slides - afterDelete.slides,
	};
}

/** The same edits through the UI: ribbon New Slide, Ctrl+M, and the inline text editor. */
async function tryUiEdits(page: Page): Promise<{ slides: number; editing: boolean }> {
	const text = slideElements(page).filter({ hasText: /\w{3}/u }).first();
	await text.dblclick({ timeout: 3000 }).catch(() =>
		// A read-only deck does not take pointer events on its elements, so aim at the slide itself.
		text.dblclick({ force: true }),
	);
	await page.waitForTimeout(500);
	const editing = (await page.locator('[data-inline-editor]').count()) > 0;
	await page.keyboard.press('Escape');
	await openRibbonTab(page, 'Home');
	const newSlide = page.getByRole('button', { name: /new slide/iu }).first();
	if (await newSlide.isEnabled().catch(() => false)) {
		await newSlide.click({ timeout: 3000 }).catch(() => undefined);
	}
	await pressShortcut(page, 'Control+m');
	// Back to the first slide, where the handle edits expect the deck's own elements.
	await page.evaluate(() => window.__pptxViewer.setActiveSlideIndex(0));
	await page.waitForTimeout(300);
	return { slides: await page.evaluate(() => window.__pptxViewer.getSlideCount()), editing };
}

async function addSlide(page: Page): Promise<void> {
	await openRibbonTab(page, 'Home');
	await page
		.getByRole('button', { name: /new slide/iu })
		.first()
		.click();
	await expect(unsaved(page).first()).toBeVisible();
}

async function fileSave(page: Page): Promise<void> {
	const backstage = await openBackstage(page);
	await backstage.getByRole('button', { name: 'Save', exact: true }).click();
}

test.describe('OneDrive under File > Open', () => {
	test('File > Save uploads to the same item and the status bar reads All saved', async ({
		page,
	}) => {
		const cloud = await routeOneDriveGraph(page, {
			mine: { name: 'Quarterly review.pptx', bytes: await deck() },
		});
		await openFromOneDrive(page, 'Quarterly review');
		await addSlide(page);
		await fileSave(page);
		// "All saved", or the autosave line once its recovery copy landed (it takes precedence).
		await expect(page.getByText(/^(All saved|Saved (just now|\d+ min.*))$/u).first()).toBeVisible({
			timeout: 30_000,
		});
		expect(cloud.writes).toStrictEqual(['mine']);
		await expect(unsaved(page)).toHaveCount(0);
		// Saving writes the saved slide identity back; moving off the new slide is no edit.
		await thumbnail(page, 1).click();
		await page.waitForTimeout(500);
		await expect(unsaved(page)).toHaveCount(0);
	});

	test('a failed upload keeps the deck unsaved', async ({ page }) => {
		const cloud = await routeOneDriveGraph(page, {
			mine: { name: 'Quarterly review.pptx', bytes: await deck() },
		});
		await openFromOneDrive(page, 'Quarterly review');
		await addSlide(page);
		cloud.failUploads(503);
		await fileSave(page);
		await page.waitForTimeout(1500);
		expect(cloud.writes).toStrictEqual([]);
		await expect(unsaved(page).first()).toBeVisible();
	});

	test('a file on another drive opens read-only: no edit from the handle or the UI, no upload', async ({
		page,
	}) => {
		const cloud = await routeOneDriveGraph(page, {
			shared: { name: 'Shared with me.pptx', bytes: await deck(), driveId: 'colleague-drive' },
		});
		await openFromOneDrive(page, 'Shared with me');
		const handle = await tryHandleEdits(page);
		// The public handle refuses every edit and leaves the deck as it was loaded.
		expect(handle.inserted).toBeUndefined();
		expect(handle).toMatchObject({
			insertedElements: 0,
			updated: false,
			duplicated: 0,
			deleted: 0,
			slidesAdded: 0,
		});
		// So do the ribbon, the keyboard and the inline text editor.
		const ui = await tryUiEdits(page);
		expect(ui.slides).toBe(handle.slides);
		expect(ui.editing).toBe(false);
		// Nothing became unsaved, and Save (the quick-access command and File > Save) uploads nothing.
		await expect(unsaved(page)).toHaveCount(0);
		await pressShortcut(page, 'Control+s');
		const backstage = await openBackstage(page);
		const save = backstage.getByRole('button', { name: 'Save', exact: true });
		if (await save.isEnabled()) {
			await save.click();
		}
		await page.waitForTimeout(1500);
		expect(cloud.writes).toStrictEqual([]);
		await expect(unsaved(page)).toHaveCount(0);
	});

	test('editing returns when an editable file is opened after a read-only one', async ({
		page,
	}) => {
		const cloud = await routeOneDriveGraph(page, {
			shared: { name: 'Shared with me.pptx', bytes: await deck(), driveId: 'colleague-drive' },
			mine: { name: 'Quarterly review.pptx', bytes: await deck() },
		});
		await openFromOneDrive(page, 'Shared with me');
		const readOnly = await tryHandleEdits(page);
		expect(readOnly.inserted).toBeUndefined();
		expect(readOnly).toMatchObject({
			insertedElements: 0,
			updated: false,
			duplicated: 0,
			deleted: 0,
			slidesAdded: 0,
		});

		await openInBackstage(page, 'Quarterly review');
		const ui = await tryUiEdits(page);
		expect(ui.editing).toBe(true);
		expect(ui.slides).toBeGreaterThan(readOnly.slides);
		const editable = await tryHandleEdits(page);
		expect(editable.inserted).toBeDefined();
		expect(editable).toMatchObject({
			insertedElements: 1,
			updated: true,
			duplicated: 1,
			deleted: 1,
			slidesAdded: 1,
		});
		await expect(unsaved(page).first()).toBeVisible();
		await fileSave(page);
		await expect(unsaved(page)).toHaveCount(0, { timeout: 30_000 });
		expect(cloud.writes).toStrictEqual(['mine']);
	});
});

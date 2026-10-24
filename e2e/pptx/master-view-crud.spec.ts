/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright spec */
import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

import { expect, test } from '@playwright/test';
import type { Locator, Page } from '@playwright/test';
import JSZip from 'jszip';
import { PptxHandler } from 'pptx-viewer-core';

import { savePptxViaBackstage } from './save-pptx';
import {
	elementWithText,
	fixture,
	loadDeck as openDeck,
	resetTabSession,
	SAMPLE_DECK,
	slideStage,
	thumbnail,
} from './support/deck';
import { downloadBytes } from './support/exports';
import { readFixtureLayout } from './support/fixture-layout';
import { dragBy } from './support/keyboard';

/**
 * Wave-4 B4: the Slide Master view sidebar's Insert / Duplicate / Delete /
 * Rename Layout commands, in every binding. The commands are ZIP surgery in
 * core and a decision list in shared (`masterViewCrudActions`); what this
 * pins per binding is the WIRING: the row exists, a click adopts the rebuilt
 * deck (the rail grows/shrinks), the fresh selection follows the command, and
 * a rename reaches the rail label.
 */

const fixturePath = resolve(
	fileURLToPath(new URL('./fixtures/master-views.pptx', import.meta.url)),
);
const RENAMED_LAYOUT = 'Wave4 Renamed Layout';

function ribbonTab(page: Page, name: string): Locator {
	return page.getByRole('tab', { name, exact: true }).first();
}

function masterTabs(page: Page): Locator {
	return page
		.getByRole('tablist')
		.filter({ has: page.getByRole('tab', { name: 'Handout', exact: true }) })
		.first();
}

/** The sidebar that owns the master tablist. */
function sidebar(page: Page): Locator {
	return masterTabs(page).locator('..');
}

function crudButton(page: Page, id: string): Locator {
	return sidebar(page).getByTestId(`pptx-master-crud-${id}`);
}

/**
 * The rail's master + layout entries plus the close control: every sidebar
 * button that is neither a tab nor a CRUD command. Only DELTAS are asserted,
 * so the constant close control never matters.
 */
function railEntries(page: Page): Locator {
	return sidebar(page).locator('button:not([role="tab"]):not([data-testid^="pptx-master-crud-"])');
}

async function loadDeck(page: Page): Promise<void> {
	await resetTabSession(page);
	await page.goto('/');
	await page.locator('#file-input').setInputFiles(fixturePath);
	await page.locator('[aria-roledescription="slide"]').first().waitFor();
}

/**
 * Where `target` sits on the slide, as a fraction of the stage width: the stage itself moves when
 * Slide Master view closes, so screen coordinates alone would not compare.
 */
async function slideOffset(page: Page, target: Locator): Promise<number> {
	const [box, stage] = await Promise.all([target.boundingBox(), slideStage(page).boundingBox()]);
	if (!box || !stage) {
		throw new Error('the element or the stage has no bounding box');
	}
	return (box.x - stage.x) / stage.width;
}

function undoButton(page: Page): Locator {
	return page.getByRole('button', { name: /^undo/iu }).first();
}

function redoButton(page: Page): Locator {
	return page.getByRole('button', { name: /^redo/iu }).first();
}

/** Slide 1's layout and a layout on its master that no slide uses, read from the deck itself. */
async function readLayoutPair(path: string): Promise<{ home: string; spare: string }> {
	const data = await new PptxHandler().load(new Uint8Array(await readFile(path)).buffer);
	const homePath = data.slides[0]?.layoutPath;
	const used = new Set(data.slides.map((slide) => slide.layoutPath));
	const layouts =
		data.slideMasters?.find((master) => master.layouts?.some((layout) => layout.path === homePath))
			?.layouts ?? [];
	const home = layouts.find((layout) => layout.path === homePath)?.name;
	const spare = layouts.find(
		(layout) => !used.has(layout.path) && layout.name && layout.name !== home,
	)?.name;
	if (!home || !spare) {
		throw new Error(`${path} has no unused layout beside the one slide 1 uses`);
	}
	return { home, spare };
}

/**
 * Open Home > Layout, the active slide's layout gallery. As in `layout-gallery-parity.spec.ts`,
 * a hover then a click opens it in every binding (Angular reveals it on hover).
 */
async function openLayoutMenu(page: Page): Promise<Locator> {
	await ribbonTab(page, 'Home').click();
	const trigger = page.getByRole('button', { name: /^layout$/iu }).first();
	await trigger.hover();
	await trigger.click().catch(() => {});
	const entries = page.locator('[data-testid="layout-gallery-menu"]:visible button');
	await expect(entries.first()).toBeVisible();
	return entries;
}

async function closeLayoutMenu(page: Page): Promise<void> {
	await page.keyboard.press('Escape');
	await page.mouse.move(0, 0);
	await expect(page.locator('[data-testid="layout-gallery-menu"]:visible button')).toHaveCount(0);
}

async function applyLayout(page: Page, name: string): Promise<void> {
	const entries = await openLayoutMenu(page);
	await entries
		.filter({ has: page.getByText(name, { exact: true }) })
		.first()
		.click();
	await closeLayoutMenu(page);
	// The apply is asynchronous in some bindings: wait until it has landed, or the next step (another
	// apply, an Undo) races it and a late apply overwrites what that step did.
	await expectCurrentLayout(page, name);
}

/** The Layout menu marks the active slide's layout, and only that one. */
async function expectCurrentLayout(page: Page, name: string): Promise<void> {
	const entries = await openLayoutMenu(page);
	const marked = entries.and(page.locator('[aria-current="true"]'));
	await expect(marked).toHaveCount(1);
	await expect(marked.filter({ has: page.getByText(name, { exact: true }) })).toHaveCount(1);
	await closeLayoutMenu(page);
}

/** Every slide-to-layout relationship in a saved package whose target part is not in it. */
async function danglingLayoutRelationships(bytes: Uint8Array): Promise<string[]> {
	const zip = await JSZip.loadAsync(bytes);
	const dangling: string[] = [];
	const rels = Object.keys(zip.files).filter((name) =>
		/^ppt\/slides\/_rels\/slide\d+\.xml\.rels$/u.test(name),
	);
	for (const name of rels) {
		const xml = (await zip.file(name)?.async('string')) ?? '';
		for (const [relationship] of xml.matchAll(/<Relationship\b[^>]*>/gu)) {
			if (!/Type="[^"]*\/slideLayout"/u.test(relationship)) {
				continue;
			}
			const target = /Target="([^"]+)"/u.exec(relationship)?.[1] ?? '';
			const part = new URL(target, 'file:///ppt/slides/').pathname.slice(1);
			if (!zip.file(part)) {
				dangling.push(`${name} -> ${target}`);
			}
		}
	}
	return dangling;
}

/** The sidebar's first button closes Slide Master view. */
async function closeMasterView(page: Page): Promise<void> {
	await sidebar(page).getByRole('button').first().click();
	await expect(masterTabs(page)).toHaveCount(0);
}

async function enterMasterView(page: Page): Promise<void> {
	await ribbonTab(page, 'View').click();
	await page
		.getByRole('toolbar')
		.first()
		.getByRole('button', { name: 'Slide Master', exact: true })
		.click();
	await expect(masterTabs(page)).toBeVisible();
}

test.describe('slide master view CRUD', () => {
	test('inserts a layout, renames it, and deletes it again', async ({ page }) => {
		await loadDeck(page);
		await enterMasterView(page);

		await expect(crudButton(page, 'addLayout')).toBeEnabled();
		const before = await railEntries(page).count();
		expect(before).toBeGreaterThan(0);

		await crudButton(page, 'addLayout').click();
		await expect(railEntries(page)).toHaveCount(before + 1);
		// The new layout is the selection, so its own commands are live.
		await expect(crudButton(page, 'deleteLayout')).toBeEnabled();
		await expect(crudButton(page, 'renameLayout')).toBeEnabled();

		page.once('dialog', (dialog) => void dialog.accept(RENAMED_LAYOUT));
		await crudButton(page, 'renameLayout').click();
		await expect(sidebar(page).getByRole('button', { name: RENAMED_LAYOUT })).toBeVisible();

		await crudButton(page, 'deleteLayout').click();
		await expect(railEntries(page)).toHaveCount(before);
		await expect(sidebar(page).getByRole('button', { name: RENAMED_LAYOUT })).toHaveCount(0);
	});

	test('keeps the slide on screen, the undo history and an unsaved edit', async ({ page }) => {
		// A command rebuilds the package through a fresh handler. Adopting the rebuilt deck as a
		// newly opened one went back to slide 1 and emptied the undo history, and a command run on
		// the deck as loaded dropped every edit made since.
		await openDeck(page, SAMPLE_DECK);
		await thumbnail(page, 2).click();
		await expect(page.getByText(/\b2 of 7\b/u).first()).toBeVisible();
		const agenda = elementWithText(page, 'Agenda');
		const start = await slideOffset(page, agenda);
		// A drag selects and moves it in one gesture (a click can start editing a placeholder's text).
		await dragBy(page, agenda, 60);
		await expect.poll(() => slideOffset(page, agenda)).toBeGreaterThan(start + 0.02);
		const moved = await slideOffset(page, agenda);
		await expect(undoButton(page)).toBeEnabled();

		await enterMasterView(page);
		const before = await railEntries(page).count();
		await crudButton(page, 'addLayout').click();
		await expect(railEntries(page)).toHaveCount(before + 1);
		await closeMasterView(page);

		await expect(page.getByText(/\b2 of 7\b/u).first()).toBeVisible();
		await expect(undoButton(page)).toBeEnabled();
		await expect.poll(() => slideOffset(page, agenda)).toBeCloseTo(moved, 2);
	});

	test('keeps the slide pictures loading', async ({ page }) => {
		// The rebuilt deck comes from a fresh handler. Taking its slides (their pictures were never
		// loaded) or disposing the handler the slides on screen came from (it revokes their Blob
		// URLs) left every picture on the slide broken.
		await openDeck(page, fixture('image-alpha-effects.pptx'));
		const pictures = slideStage(page).locator('img');
		await expect(pictures.first()).toBeVisible();
		const count = await pictures.count();
		const broken = () =>
			pictures.evaluateAll(
				(images: HTMLImageElement[]) =>
					images.filter((image) => !image.complete || image.naturalWidth === 0).length,
			);
		await expect.poll(broken).toBe(0);

		await enterMasterView(page);
		const before = await railEntries(page).count();
		await crudButton(page, 'addLayout').click();
		await expect(railEntries(page)).toHaveCount(before + 1);
		await closeMasterView(page);

		await expect(pictures).toHaveCount(count);
		await expect.poll(broken).toBe(0);
	});

	test('leaves no undo or redo step that puts a slide back on a deleted layout', async ({
		page,
	}) => {
		// Delete Layout only checks the slides as they are now. A slide on the layout in an older
		// undo step, or in an undone redo step, went back onto it, and the next save wrote a slide
		// relationship to a part the package no longer has.
		const { home, spare } = await readLayoutPair(SAMPLE_DECK);
		await openDeck(page, SAMPLE_DECK);
		// Slide 1 goes onto the spare layout and back, so an undo step has it on the spare layout.
		await applyLayout(page, spare);
		await applyLayout(page, home);

		await enterMasterView(page);
		await sidebar(page).getByRole('button', { name: spare, exact: true }).click();
		await expect(crudButton(page, 'deleteLayout')).toBeEnabled();
		const before = await railEntries(page).count();
		await crudButton(page, 'deleteLayout').click();
		await expect(railEntries(page)).toHaveCount(before - 1);
		await closeMasterView(page);

		// How many steps are left differs by binding (some also restore the masters, so a delete
		// leaves none); none of them may reach the deleted layout.
		for (let step = 0; step < 6 && (await undoButton(page).isEnabled()); step++) {
			await undoButton(page).click();
			await expectCurrentLayout(page, home);
		}
		for (let step = 0; step < 6 && (await redoButton(page).isEnabled()); step++) {
			await redoButton(page).click();
			await expectCurrentLayout(page, home);
		}
		const saved = await downloadBytes(await savePptxViaBackstage(page));
		expect(await danglingLayoutRelationships(saved)).toStrictEqual([]);
	});

	test('a layout a slide uses cannot be deleted', async ({ page }) => {
		await loadDeck(page);
		await enterMasterView(page);

		const layout = await readFixtureLayout(fixturePath);
		await sidebar(page).getByRole('button', { name: layout.name, exact: true }).click();
		await expect(crudButton(page, 'duplicateLayout')).toBeEnabled();
		await expect(crudButton(page, 'deleteLayout')).toBeDisabled();
		await expect(crudButton(page, 'deleteLayout')).toHaveAttribute('title', /in use|used by/iu);
	});
});

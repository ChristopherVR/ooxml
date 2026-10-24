/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright spec */
import { resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

import { expect, test } from '@playwright/test';
import type { Locator, Page } from '@playwright/test';

import {
	elementWithText,
	fixture,
	loadDeck as openDeck,
	resetTabSession,
	SAMPLE_DECK,
	slideStage,
	thumbnail,
} from './support/deck';
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

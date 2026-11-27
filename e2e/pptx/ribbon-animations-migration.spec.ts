import { writeFile } from 'node:fs/promises';

/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright API */
import { expect, test } from '@playwright/test';
import type { Page } from '@playwright/test';
import JSZip from 'jszip';
import { PptxAnimationWriteService, PresentationBuilder } from 'pptx-viewer-core';

import { savePptxViaBackstage } from './save-pptx';
import {
	elementWithText,
	fixture,
	inspector,
	loadDeck,
	ribbon,
	ribbonTab,
	SAMPLE_DECK,
	selectElement,
} from './support/deck';
import { downloadBytes } from './support/exports';
import { chooseSelectValue, expectSelectValue } from './support/select-control';

test.use({ viewport: { width: 1440, height: 900 } });

const DECK = fixture('ribbon-galleries.pptx');
const SHAPE_TEXT = 'GALLERY SHAPE';
const GROUPS = [
	'animations.preview',
	'animations.animation',
	'animations.motionPath',
	'animations.advancedAnimation',
	'animations.timing',
];
const CONTROLS = [
	'animations.preview.preview',
	'animations.animation.gallery',
	'animations.animation.effectOptions',
	'animations.motionPath.gallery',
	'animations.advancedAnimation.addAnimation',
	'animations.advancedAnimation.animationPane',
	'animations.advancedAnimation.trigger',
	'animations.advancedAnimation.animationPainter',
	'animations.advancedAnimation.remove',
	'animations.timing.start',
	'animations.timing.duration',
];

async function openAnimations(page: Page, select = false, deck = DECK) {
	await loadDeck(page, deck);
	if (select) {
		await selectElement(page, elementWithText(page, SHAPE_TEXT));
	}
	await ribbonTab(page, 'Animations').click();
	return page.locator('pptx-ui-ribbon-animations');
}

async function savedSlideXml(page: Page): Promise<{ path: string; xml: string }> {
	const bytes = await downloadBytes(await savePptxViaBackstage(page));
	const zip = await JSZip.loadAsync(bytes);
	const xml = await zip.file('ppt/slides/slide1.xml')!.async('string');
	return { path: 'ppt/slides/slide1.xml', xml };
}

test('the Animation Pane selects which imported effect receives a Timing edit', async ({
	page,
}) => {
	await loadDeck(page, fixture('native-animation-timing.pptx'));
	await selectElement(page, elementWithText(page, 'Native animation timing'));
	await ribbonTab(page, 'Animations').click();
	const animationPane = page.locator(
		'[data-ribbon-control="animations.advancedAnimation.animationPane"] button',
	);
	const pane = inspector(page);
	const rows = pane.locator('button[data-native-effect-id]');
	const effect = pane.locator('button[data-native-effect-id="8"]');
	if (!(await effect.isVisible())) await animationPane.click();
	await expect(rows).toHaveCount(4);
	await expect(effect).toHaveAccessibleName(/Rectangle 1.*Fade/u);
	await effect.press('Enter');
	await expect(effect).toHaveAttribute('aria-pressed', 'true');
	await page.emulateMedia({ forcedColors: 'active' });
	await expect(effect).toHaveCSS('forced-color-adjust', 'none');
	await page.emulateMedia({ forcedColors: 'none' });
	const start = page
		.locator('[data-ribbon-control="animations.timing.start"]')
		.getByRole('combobox');
	await expectSelectValue(start, 'withPrevious');
	// An imported With Previous effect can now leave its click group (it is split off), so no
	// Start value is refused and the field carries no refusal reason.
	await expect(start).toHaveAccessibleDescription('');
	await start.click();
	await expect(page.getByRole('option', { name: 'On Click', exact: true })).toBeEnabled();
	await start.press('Escape');
	const duration = page.locator('[data-ribbon-control="animations.timing.duration"]');
	await expect(duration).toHaveValue('1');
	await duration.fill('1.75');
	await duration.press('Tab');
	await expect(effect).toHaveAttribute('aria-pressed', 'false');
	await effect.click();
	await expect(duration).toHaveValue('1.75');
	const delay = page.locator('[data-ribbon-control="animations.timing.delay"]');
	await delay.fill('0.35');
	await delay.press('Tab');
	await expect(effect).toHaveAttribute('aria-pressed', 'false');
	await effect.click();
	await expect(delay).toHaveValue('0.35');
	await expectSelectValue(start, 'withPrevious');
	await page.getByRole('button', { name: 'Undo', exact: true }).first().click();
	await expect(effect).toHaveAttribute('aria-pressed', 'false');
	await effect.click();
	await expect(delay).toHaveValue('0.2');
	await page.getByRole('button', { name: 'Redo', exact: true }).first().click();
	const { xml } = await savedSlideXml(page);
	expect(xml).toMatch(
		/<p:cTn\b(?=[^>]*id="8")(?=[^>]*dur="1750")(?=[^>]*nodeType="withEffect")[^>]*><p:stCondLst><p:cond delay="350"/u,
	);
	expect(xml).not.toMatch(/<p:cTn\b(?=[^>]*id="5")(?=[^>]*dur="1750")[^>]*>/u);
});

test('the Animation Pane offers imported effects without a canvas selection', async ({ page }) => {
	const animations = await openAnimations(page, false, fixture('native-animation-timing.pptx'));
	const row = inspector(page).locator('button[data-native-effect-id="8"]');
	if (!(await row.isVisible())) {
		await animations
			.locator('[data-ribbon-control="animations.advancedAnimation.animationPane"] button')
			.click();
	}
	await expect(row).toHaveAccessibleName(/Rectangle 1.*Fade/u);
	await row.click();
	await expect(
		animations.locator('[data-ribbon-control="animations.timing.duration"]'),
	).toHaveValue('1');
});

test('Faded Zoom offers the two centers and saves the selected subtype', async ({ page }, info) => {
	const { createSlide, data, handler } = await PresentationBuilder.create();
	const slide = createSlide('Blank')
		.addText('Zoom target', { x: 50, y: 50, width: 200, height: 40 })
		.build();
	slide.elements[0]!.shapeId = '2';
	slide.rawTiming = new PptxAnimationWriteService().buildTimingXml(
		[{ elementId: '2', entrance: 'entr.53', direction: 'objectCenter' }],
		undefined,
	);
	delete slide.rawTiming?.['p:extLst'];
	data.slides.push(slide);
	const path = info.outputPath('faded-zoom.pptx');
	await writeFile(path, await handler.save(data.slides));
	await openAnimations(page, false, path);
	const row = inspector(page).getByRole('button', { name: /Faded Zoom/u });
	if (!(await row.isVisible())) {
		await page
			.locator('[data-ribbon-control="animations.advancedAnimation.animationPane"] button')
			.click();
	}
	await row.click();
	const direction = page.getByRole('combobox', { name: 'Direction', exact: true });
	await expect(direction).toContainText('Object Center');
	await chooseSelectValue(page, direction, 'slideCenter');
	expect((await savedSlideXml(page)).xml).toMatch(
		/presetID="53" presetClass="entr" presetSubtype="528"/u,
	);
});

test('shared Animations ribbon exposes every canonical group and control id once', async ({
	page,
}, info) => {
	// The shared demo deck matches the recorded baseline captures.
	const animations = await openAnimations(page, false, SAMPLE_DECK);
	await page.screenshot({ path: info.outputPath('animations.png') });
	for (const id of GROUPS) {
		await expect(animations.locator(`[data-ribbon-group="${id}"]`)).toHaveCount(1);
	}
	for (const id of CONTROLS) {
		await expect(animations.locator(`[data-ribbon-control="${id}"]`)).toHaveCount(1);
	}
	// Without a selection every authoring control is disabled; the pane is not.
	for (const id of [
		'animations.preview.preview',
		'animations.advancedAnimation.addAnimation',
		'animations.advancedAnimation.remove',
		'animations.advancedAnimation.animationPainter',
	]) {
		await expect(animations.locator(`[data-ribbon-control="${id}"] button`)).toBeDisabled();
	}
	await expect(
		animations.locator('[data-ribbon-control="animations.advancedAnimation.animationPane"] button'),
	).toBeEnabled();
	await expect(animations.getByRole('button', { name: 'Fly In', exact: true })).toBeDisabled();
	await expect(
		animations.locator('[data-ribbon-control="animations.timing.duration"]'),
	).toBeDisabled();
});

async function addFlyIn(page: Page) {
	const animations = await openAnimations(page, true);
	const flyIn = animations.getByRole('button', { name: 'Fly In', exact: true });
	await expect(flyIn).toBeEnabled();
	await flyIn.focus();
	await page.keyboard.press('Space');
	return animations;
}

test('adding an effect from the gallery reaches the deck and the Animation Pane', async ({
	page,
}) => {
	const animations = await addFlyIn(page);
	// The ribbon never leaves the selection: Preview becomes usable.
	await expect(
		animations.locator('[data-ribbon-control="animations.preview.preview"] button'),
	).toBeEnabled();
	const pane = animations.locator(
		'[data-ribbon-control="animations.advancedAnimation.animationPane"] button',
	);
	if (!(await inspector(page).isVisible())) {
		await pane.click();
	}
	await expect(inspector(page)).toBeVisible();
	await expect(pane).toHaveAttribute('aria-pressed', 'true');
	const { xml } = await savedSlideXml(page);
	expect(xml).toContain('<p:timing');
	expect(xml).toMatch(/presetClass="entr"/u);
	expect(xml).toMatch(/presetID="2"/u);
});

test('the selected Fade effect has timing controls and no Direction picker', async ({ page }) => {
	const animations = await openAnimations(page, true);
	await animations.getByRole('button', { name: 'Fade In', exact: true }).click();
	await expect(
		animations.locator('[data-ribbon-control="animations.timing.duration"]'),
	).toBeEnabled();
	await expect(ribbon(page).getByRole('combobox', { name: 'Direction', exact: true })).toHaveCount(
		0,
	);
});

test('undo removes the added effect from the saved deck', async ({ page }) => {
	await addFlyIn(page);
	const undo = page.getByRole('button', { name: 'Undo', exact: true }).first();
	await expect(undo).toBeEnabled();
	await undo.click();
	// Saving while the undo is still in flight loses it, and every poll would save again:
	// wait for the UI to show the undo (Redo becomes available), then save once.
	await expect(page.getByRole('button', { name: 'Redo', exact: true }).first()).toBeEnabled();
	expect((await savedSlideXml(page)).xml).not.toContain('<p:timing');
});

test('a redone effect survives save and reload and previews on the saved deck', async ({
	page,
}, info) => {
	await addFlyIn(page);
	await page.getByRole('button', { name: 'Undo', exact: true }).first().click();
	await page.getByRole('button', { name: 'Redo', exact: true }).first().click();
	const saved = info.outputPath('animation-edit.pptx');
	const bytes = await downloadBytes(await savePptxViaBackstage(page));
	await writeFile(saved, bytes);
	const zip = await JSZip.loadAsync(bytes);
	expect(await zip.file('ppt/slides/slide1.xml')!.async('string')).toMatch(/presetID="2"/u);

	await loadDeck(page, saved);
	await selectElement(page, elementWithText(page, SHAPE_TEXT));
	await ribbonTab(page, 'Animations').click();
	// The saved effect is the selection's own: Preview starts a real CSS animation.
	await page.evaluate(() => {
		document.documentElement.removeAttribute('data-ribbon-animation-preview');
		const record = (event: AnimationEvent) => {
			document.documentElement.setAttribute('data-ribbon-animation-preview', event.animationName);
			document.removeEventListener('animationstart', record, true);
		};
		document.addEventListener('animationstart', record, true);
	});
	await page
		.locator('[data-ribbon-control="animations.preview.preview"]')
		.getByRole('button')
		.click();
	await expect(page.locator('html')).toHaveAttribute('data-ribbon-animation-preview', /^pptx-/u);
});

test('a motion path from the gallery is written to the deck', async ({ page }) => {
	const animations = await openAnimations(page, true);
	await animations
		.locator('[data-ribbon-control="animations.motionPath.gallery"] button')
		.first()
		.click();
	const { xml } = await savedSlideXml(page);
	expect(xml).toContain('<p:animMotion');
});

test('Animations controls retain public customization IDs', async ({ page }) => {
	const customization = {
		ribbon: {
			hiddenButtons: ['animations.advancedAnimation.addAnimation', 'animations.timing.start'],
			hiddenGroups: ['animations.motionPath'],
		},
	};
	await loadDeck(
		page,
		DECK,
		`/?customization=${encodeURIComponent(JSON.stringify(customization))}`,
	);
	await ribbonTab(page, 'Animations').click();
	await expect(
		page.locator('[data-ribbon-control="animations.advancedAnimation.addAnimation"]'),
	).toBeHidden();
	await expect(page.locator('[data-ribbon-control="animations.timing.start"]')).toBeHidden();
	await expect(page.locator('[data-ribbon-group="animations.motionPath"]')).toBeHidden();
	await expect(page.locator('[data-ribbon-control="animations.timing.duration"]')).toBeVisible();
	await expect(page.locator('[data-ribbon-group="animations.animation"]')).toBeVisible();
});

test.describe('touch Animations controls', () => {
	test.use({ hasTouch: true });
	test('targets, focus, theme tokens and forced colors remain usable', async ({ page }) => {
		// No canvas selection: the Animation Pane is the one command that never needs one.
		const animations = await openAnimations(page);
		await page.setViewportSize({ width: 900, height: 900 });
		const flyIn = animations.getByRole('button', { name: 'Fly In', exact: true });
		const box = await flyIn.boundingBox();
		expect(box!.height).toBeGreaterThanOrEqual(44);
		const pane = animations.locator(
			'[data-ribbon-control="animations.advancedAnimation.animationPane"] button',
		);
		const paneBox = await pane.boundingBox();
		expect(paneBox!.height).toBeGreaterThanOrEqual(44);
		await pane.focus();
		// Keyboard modality makes the focus ring visible.
		await page.keyboard.press('Shift+Tab');
		await page.keyboard.press('Tab');
		await expect(pane).toBeFocused();
		await expect(pane).toHaveCSS('outline-style', 'solid');
		await animations.evaluate((host) => host.style.setProperty('--pptx-primary', '#123456'));
		if ((await pane.getAttribute('aria-pressed')) !== 'true') {
			await pane.tap();
		}
		await expect(pane).toHaveAttribute('aria-pressed', 'true');
		await expect(pane).toHaveCSS('color', 'rgb(18, 52, 86)');
		await page.emulateMedia({ forcedColors: 'active' });
		await expect(flyIn).toBeVisible();
		await expect(
			animations.locator('[data-ribbon-control="animations.animation.gallery"]'),
		).toHaveCSS('border-top-style', 'solid');
	});
});

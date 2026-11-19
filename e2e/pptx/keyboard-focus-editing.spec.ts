/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright spec, `test`/`expect` come from @playwright/test */
/**
 * Keyboard focus survives editing, in every binding: a slide change, the notes
 * field, text editing, a dialog and a mouse press never strand it on the page
 * body or trap it, and Enter edits the selected object's text as in PowerPoint.
 */
import { expect, test } from '@playwright/test';
import type { Locator, Page } from '@playwright/test';

import { fixture, loadDeck, ribbonTab, SAMPLE_DECK, slideElements } from './support/deck';

test.use({ viewport: { width: 1440, height: 900 } });

/** A one-slide deck holding exactly two shapes. */
const TWO_SHAPES = fixture('format-painter.pptx');

const viewport = (page: Page): Locator => page.locator('[data-pptx-viewport]').first();
const inlineEditor = (page: Page): Locator => page.locator('[data-inline-editor]');
const notesField = (page: Page): Locator =>
	page.locator('#slide-notes-content [contenteditable="true"]').first();

/** Where keyboard focus is: the slide itself, inside it, a named pane, or nowhere. */
function focusPlace(page: Page): Promise<string> {
	return page.evaluate(() => {
		let active = document.activeElement;
		while (active?.shadowRoot?.activeElement) active = active.shadowRoot.activeElement;
		if (!active || active === document.body) return 'body';
		/** The nearest ancestor matching `selector`, across shadow roots. */
		const up = (selector: string): Element | null => {
			for (let node: Node | null = active; node;) {
				if (node instanceof Element && node.matches(selector)) return node;
				node = node.parentNode ?? (node instanceof ShadowRoot ? node.host : null);
			}
			return null;
		};
		// A dialog can be mounted inside the slide viewport, so it is asked first.
		if (up('[role="dialog"], dialog')) return 'dialog';
		const slide = up('[data-pptx-viewport]');
		if (slide) return slide === active ? 'slide' : 'in-slide';
		if (
			active.closest('#slide-notes-content') &&
			active.matches('[contenteditable="true"], textarea')
		)
			return 'notes-field';
		return `${active.tagName}:${(active.getAttribute('aria-label') ?? '').slice(0, 30)}`;
	});
}

async function focusSlide(page: Page): Promise<void> {
	await viewport(page).focus();
	expect(await focusPlace(page)).toBe('slide');
}

async function openNotes(page: Page): Promise<void> {
	await page
		.getByRole('button', { name: /^toggle notes$/iu })
		.first()
		.focus();
	await page.keyboard.press('Enter');
	await expect(notesField(page)).toBeVisible();
	await notesField(page).click();
	await page.keyboard.type('n');
	expect(await focusPlace(page)).toBe('notes-field');
}

test.describe('keyboard focus while editing', () => {
	test('a slide change keeps focus on the slide', async ({ page }) => {
		await loadDeck(page, SAMPLE_DECK);
		await focusSlide(page);
		await page.keyboard.press('PageDown');
		await page.waitForTimeout(300);
		expect(await focusPlace(page), 'PageDown must not drop focus').toBe('slide');
		await page.keyboard.press('PageDown');
		expect(await focusPlace(page)).toBe('slide');
		await page.keyboard.press('Tab');
		expect(await focusPlace(page), 'the slide is still a Tab stop').not.toBe('body');
	});

	test('the notes field lets Tab, Shift+Tab, F6 and Escape leave', async ({ page }) => {
		await loadDeck(page, TWO_SHAPES);
		await openNotes(page);
		await page.keyboard.press('Tab');
		expect(await focusPlace(page), 'Tab').not.toMatch(/^(notes-field|body)$/u);
		await notesField(page).click();
		await page.keyboard.press('Shift+Tab');
		expect(await focusPlace(page), 'Shift+Tab').not.toMatch(/^(notes-field|body)$/u);
		await notesField(page).click();
		await page.keyboard.press('F6');
		expect(await focusPlace(page), 'F6').not.toMatch(/^(notes-field|body)$/u);
		await notesField(page).click();
		await page.keyboard.press('Shift+F6');
		expect(await focusPlace(page), 'Shift+F6').not.toMatch(/^(notes-field|body)$/u);
		await notesField(page).click();
		await page.keyboard.press('Escape');
		expect(await focusPlace(page), 'Escape goes back to the slide').toBe('slide');
	});

	test('Enter on a selected object edits its text, and Escape returns to the slide', async ({
		page,
	}) => {
		await loadDeck(page, TWO_SHAPES);
		await focusSlide(page);
		await page.keyboard.press('Enter');
		await page.keyboard.press('Enter');
		await expect(inlineEditor(page), 'Enter must start text editing').toHaveCount(1);
		await page.keyboard.type('QX');
		await page.keyboard.press('Escape');
		await expect(inlineEditor(page)).toHaveCount(0);
		expect(await focusPlace(page), 'Escape must return focus to the slide').toBe('slide');
		await expect(slideElements(page).filter({ hasText: 'QX' })).toHaveCount(1);
		// Still inside the slide: Tab walks the objects and Delete removes one.
		await page.keyboard.press('Tab');
		expect(await focusPlace(page)).toBe('slide');
		await page.keyboard.press('Delete');
		await expect(slideElements(page)).toHaveCount(1);
	});

	test('Escape after a double-click edit returns focus to the slide', async ({ page }) => {
		await loadDeck(page, TWO_SHAPES);
		await slideElements(page).first().dblclick();
		await expect(inlineEditor(page)).toHaveCount(1);
		await page.keyboard.press('Escape');
		await expect(inlineEditor(page)).toHaveCount(0);
		expect(await focusPlace(page)).toBe('slide');
	});

	test('the hyperlink dialog keeps the keyboard and gives it back', async ({ page }) => {
		await loadDeck(page, TWO_SHAPES);
		await focusSlide(page);
		await page.keyboard.press('Enter');
		await page.keyboard.press('Control+k');
		await expect(page.getByRole('dialog')).toBeVisible();
		await expect.poll(() => focusPlace(page), 'focus goes into the dialog').toBe('dialog');
		for (let press = 0; press < 10; press += 1) {
			await page.keyboard.press('Tab');
			expect(await focusPlace(page), `Tab ${press + 1}`).toBe('dialog');
		}
		await page.keyboard.press('Shift+Tab');
		expect(await focusPlace(page)).toBe('dialog');
		await page.keyboard.press('Delete');
		await expect(slideElements(page), 'Delete in a dialog keeps the object').toHaveCount(2);
		await page.keyboard.press('Escape');
		await expect(page.getByRole('dialog')).toHaveCount(0);
		expect(await focusPlace(page), 'focus returns to the slide').toBe('slide');
	});

	test('the File tab takes focus into the backstage and Escape brings it back', async ({
		page,
	}) => {
		await loadDeck(page, TWO_SHAPES);
		await ribbonTab(page, 'File').focus();
		await page.keyboard.press('Enter');
		await expect(page.getByRole('dialog')).toBeVisible();
		await expect.poll(() => focusPlace(page), 'focus goes into the backstage').toBe('dialog');
		await page.keyboard.press('Tab');
		expect(await focusPlace(page)).toBe('dialog');
		await page.keyboard.press('Escape');
		await expect(page.getByRole('dialog')).toHaveCount(0);
		expect(await focusPlace(page), 'focus returns to the ribbon').not.toBe('body');
	});

	test('the slide focus ring shows for the keyboard, not for a click', async ({ page }) => {
		await loadDeck(page, TWO_SHAPES);
		const ring = () => viewport(page).evaluate((node) => getComputedStyle(node).outlineStyle);
		await slideElements(page).first().click();
		expect(await focusPlace(page)).toMatch(/slide$/u);
		expect(await ring(), 'no ring after a click').toBe('none');
		// F6 round the panes back to the slide: the keyboard put focus there, so it shows.
		for (let press = 0; press < 8; press += 1) {
			await page.keyboard.press('F6');
			if ((await focusPlace(page)) === 'slide') break;
		}
		expect(await focusPlace(page)).toBe('slide');
		expect(await ring(), 'a ring once the keyboard is used').not.toBe('none');
	});
});

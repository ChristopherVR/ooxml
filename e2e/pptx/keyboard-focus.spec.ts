/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright spec, `test`/`expect` come from @playwright/test */
/**
 * The keyboard can always reach the slide, go into it and leave it, in every binding.
 *
 * Tab order only exists in a real browser, so this is where it is proved: the
 * slide is one Tab stop, Enter goes in, Tab walks the objects as in PowerPoint,
 * Escape and F6 come back out.
 */
import { expect, test } from '@playwright/test';
import type { Page } from '@playwright/test';

import { fixture, loadDeck, slideElements } from './support/deck';

test.use({ viewport: { width: 1440, height: 900 } });

/** A one-slide deck holding exactly two shapes. */
const TWO_SHAPES = fixture('format-painter.pptx');

const viewport = (page: Page) => page.locator('[data-pptx-viewport]').first();

/** Where keyboard focus is: the slide, somewhere else, or nowhere. */
function focusPlace(page: Page): Promise<string> {
	return page.evaluate(() => {
		let active = document.activeElement;
		while (active?.shadowRoot?.activeElement) active = active.shadowRoot.activeElement;
		if (!active || active === document.body) return 'body';
		if (active.closest('[data-pptx-viewport]')) return 'slide';
		const label = active.getAttribute('aria-label') ?? active.textContent ?? '';
		return `${active.tagName}:${label.trim().slice(0, 40)}`;
	});
}

/** Tab until focus has left the slide; links on the slide are stops of their own. */
async function tabOutOfSlide(page: Page): Promise<string> {
	for (let press = 0; press < 12; press += 1) {
		await page.keyboard.press('Tab');
		const place = await focusPlace(page);
		if (place !== 'slide') return place;
	}
	return 'slide';
}

async function focusSlide(page: Page): Promise<void> {
	await viewport(page).focus();
	expect(await focusPlace(page)).toBe('slide');
}

test.describe('keyboard focus on the slide', () => {
	test('the slide is one Tab stop that lets the keyboard move on', async ({ page }) => {
		await loadDeck(page, TWO_SHAPES);
		await focusSlide(page);
		const next = await tabOutOfSlide(page);
		expect(next, 'Tab must leave the slide').not.toBe('slide');
		expect(next).not.toBe('body');
	});

	test('no Tab stop repeats: the whole page can be walked', async ({ page }) => {
		await loadDeck(page, TWO_SHAPES);
		let previous = '';
		let repeats = 0;
		let reachedSlide = false;
		for (let press = 0; press < 400 && !reachedSlide; press += 1) {
			await page.keyboard.press('Tab');
			const place = await focusPlace(page);
			reachedSlide = place === 'slide';
			repeats = place === previous ? repeats + 1 : 0;
			previous = place;
			expect(repeats, `stuck on ${place}`).toBeLessThan(3);
		}
		expect(reachedSlide, 'Tab must reach the slide').toBe(true);
		expect(await tabOutOfSlide(page)).not.toBe('slide');
	});

	test('Enter goes in, Tab walks the objects, Escape comes back out', async ({ page }) => {
		await loadDeck(page, TWO_SHAPES);
		await focusSlide(page);
		await page.keyboard.press('Enter');
		await page.keyboard.press('Delete');
		await expect(slideElements(page), 'Enter must select the first object').toHaveCount(1);
		const survivorOfEnter = await slideElements(page).first().textContent();

		await loadDeck(page, TWO_SHAPES);
		await focusSlide(page);
		await page.keyboard.press('Enter');
		await page.keyboard.press('Tab');
		expect(await focusPlace(page), 'Tab inside the slide keeps focus there').toBe('slide');
		await page.keyboard.press('Delete');
		await expect(slideElements(page)).toHaveCount(1);
		expect(await slideElements(page).first().textContent()).not.toBe(survivorOfEnter);

		await loadDeck(page, TWO_SHAPES);
		await focusSlide(page);
		await page.keyboard.press('Enter');
		await page.keyboard.press('Escape');
		await page.keyboard.press('Escape');
		expect(await tabOutOfSlide(page), 'after Escape, Tab moves on again').not.toBe('slide');
	});

	test('a click goes in, so Tab walks from the clicked object', async ({ page }) => {
		await loadDeck(page, TWO_SHAPES);
		await slideElements(page).first().click();
		await page.keyboard.press('Tab');
		expect(await focusPlace(page)).toBe('slide');
	});

	test('F6 moves between the panes and back to the slide', async ({ page }) => {
		await loadDeck(page, TWO_SHAPES);
		await focusSlide(page);
		const visited: string[] = [];
		for (let press = 0; press < 8; press += 1) {
			await page.keyboard.press('F6');
			const place = await focusPlace(page);
			visited.push(place);
			if (place === 'slide') break;
		}
		expect(visited.at(-1), `F6 order: ${visited.join(' > ')}`).toBe('slide');
		expect(visited.length, 'F6 must visit at least one other pane').toBeGreaterThan(1);
		expect(visited).not.toContain('body');

		await page.keyboard.press('Shift+F6');
		expect(await focusPlace(page)).toBe(visited.at(-2));
		await page.keyboard.press('F6');
		expect(await focusPlace(page)).toBe('slide');
		// F6 landed on the slide, so Tab now walks its objects, as in PowerPoint.
		await page.keyboard.press('Tab');
		expect(await focusPlace(page)).toBe('slide');
	});
});

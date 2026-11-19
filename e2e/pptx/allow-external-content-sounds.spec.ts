/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright spec, `test`/`expect` come from @playwright/test */
/**
 * Linked sounds are linked content too: the slide show requests none of a
 * deck's linked transition, effect or hyperlink sounds before the reader
 * chooses Enable Content (or with the host's `allowExternalContent: false`),
 * and plays them once the reader has. The deck links every sound to an
 * external origin, answered here with `page.route`, so every request the show
 * causes is recorded.
 */
import { expect, test } from '@playwright/test';
import type { Page } from '@playwright/test';

import { resetTabSession } from './support/deck';
import { EXTERNAL_ORIGIN, LINKED_SOUNDS, linkedSoundDeck } from './support/external-content-deck';

/** Record every request to the linked origin, answered offline. */
async function recordExternalRequests(page: Page): Promise<string[]> {
	const paths: string[] = [];
	await page.route(`${EXTERNAL_ORIGIN}/**`, (route) => {
		paths.push(new URL(route.request().url()).pathname);
		return route.fulfill({ status: 404, body: '' });
	});
	return paths;
}

async function openSoundDeck(page: Page, query = ''): Promise<void> {
	await resetTabSession(page);
	await page.goto(`/${query}`);
	await page.locator('#file-input').setInputFiles({
		name: 'linked-sounds.pptx',
		mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
		buffer: await linkedSoundDeck(),
	});
	await page.locator('[data-element-id]').filter({ hasText: 'Linked sounds' }).first().waitFor();
}

/** Start the show, move to slide 2 (its transition sound), play its effect, then click the action shape. */
async function playEverySound(page: Page): Promise<void> {
	const slideShow = page.getByRole('button', { name: /^slide show$/iu });
	await (
		(await slideShow.count()) > 0
			? slideShow.last()
			: page.getByRole('button', { name: /present/iu }).first()
	).click();
	await page.waitForTimeout(700);
	await page.keyboard.press('PageDown');
	await page.waitForTimeout(900);
	await page.keyboard.press('PageDown');
	await page.waitForTimeout(600);
	const action = page.getByText('Action sound', { exact: true }).last();
	const box = await action.boundingBox();
	if (box) {
		await page.mouse.click(box.x + box.width / 2, box.y + box.height / 2);
	}
	await page.waitForTimeout(600);
}

test.describe('linked sounds in the slide show', () => {
	test('are not requested before Enable Content', async ({ page }) => {
		const requests = await recordExternalRequests(page);
		await openSoundDeck(page);
		// The sounds alone are linked content: the bar offers to enable them.
		await expect(page.locator('pptx-ui-external-content-bar')).toBeVisible();
		await playEverySound(page);
		expect(requests).toStrictEqual([]);
	});

	test('are not requested when the host blocks linked content', async ({ page }) => {
		const requests = await recordExternalRequests(page);
		await openSoundDeck(page, '?allowExternalContent=0');
		await playEverySound(page);
		expect(requests).toStrictEqual([]);
	});

	test('play after Enable Content, the transition sound included', async ({ page }) => {
		const requests = await recordExternalRequests(page);
		await openSoundDeck(page);
		await page
			.locator('pptx-ui-external-content-bar')
			.getByRole('button', { name: 'Enable Content' })
			.click();
		await playEverySound(page);
		await expect.poll(() => requests).toContain(LINKED_SOUNDS.transition);
		await expect.poll(() => requests).toContain(LINKED_SOUNDS.effect);
	});
});

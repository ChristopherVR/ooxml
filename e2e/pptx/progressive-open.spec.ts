/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright browser benchmark */
import type { Page } from '@playwright/test';
import { expect, test } from '@playwright/test';

import {
	makePptxOpeningDeck,
	PPTX_OPEN_LARGE,
	PPTX_OPEN_SMALL,
} from '../../scripts/perf-budget-pptx.mjs';
import { thumbnail, viewport } from './support/deck';

const PREVIEW = '[data-pptx-opening-preview]';
const CANVAS_ELEMENT = '[data-pptx-viewport] [data-element-id]';
const INLINE_EDITOR = '[data-inline-editor]';

test.describe.configure({ timeout: 90_000 });

async function upload(page: Page, size: number): Promise<void> {
	const bytes = await makePptxOpeningDeck(size);
	await page.locator('#file-input').setInputFiles({
		name: `opening-${size}.pptx`,
		mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
		buffer: Buffer.from(bytes),
	});
}

/** Measure from the input event through a frame that presented the expected result. */
async function armPaintMeasurement(
	page: Page,
	name: string,
	eventName: string,
	selector: string,
	text = '',
): Promise<void> {
	await page.evaluate(
		({ name, eventName, selector, text }) => {
			document.addEventListener(
				eventName,
				() => {
					const start = performance.now();
					const poll = () => {
						const ready = Array.from(document.querySelectorAll<HTMLElement>(selector)).some(
							(node) => {
								const content = node instanceof HTMLTextAreaElement ? node.value : node.textContent;
								return node.getClientRects().length > 0 && (!text || content?.includes(text));
							},
						);
						requestAnimationFrame(
							ready ? () => performance.measure(name, { start, end: performance.now() }) : poll,
						);
					};
					requestAnimationFrame(poll);
				},
				{ capture: true, once: true },
			);
		},
		{ name, eventName, selector, text },
	);
}

for (const size of [PPTX_OPEN_SMALL, PPTX_OPEN_LARGE]) {
	test(`opening and interaction timings for ${size} slides`, async ({ page }, info) => {
		const errors: string[] = [];
		page.on('pageerror', (error) => errors.push(error.message));
		await page.goto('/');
		await armPaintMeasurement(
			page,
			'first-slide',
			'change',
			`${PREVIEW}, ${CANVAS_ELEMENT}`,
			'Deck slide 1 shape 1',
		);
		await upload(page, size);
		const title = page
			.locator(CANVAS_ELEMENT)
			.getByText('Deck slide 1 shape 1', { exact: true })
			.first();
		await expect(title).toBeVisible({ timeout: 45_000 });
		await armPaintMeasurement(page, 'select', 'pointerdown', '[data-pptx-handle-kind="resize"]');
		await title.click();
		await expect(
			viewport(page)
				.getByRole('button', { name: /^resize /iu })
				.first(),
		).toBeVisible();
		await title.dblclick();
		const editor = page.locator(INLINE_EDITOR);
		await expect(editor).toBeVisible();
		await editor.press('End');
		await armPaintMeasurement(page, 'type', 'keydown', INLINE_EDITOR, 'shape 1Z');
		await editor.press('Z');
		await expect
			.poll(() => page.evaluate(() => performance.getEntriesByName('type', 'measure').length))
			.toBe(1);
		// Commit the edit before measuring navigation, which can rebuild the slide rail on blur.
		await editor.blur();
		await expect(editor).toBeHidden();
		await armPaintMeasurement(
			page,
			'switch',
			'pointerdown',
			CANVAS_ELEMENT,
			'Deck slide 2 shape 1',
		);
		await thumbnail(page, 2).first().click();
		await expect(
			page.locator(CANVAS_ELEMENT).getByText('Deck slide 2 shape 1', { exact: true }).first(),
		).toBeVisible();
		await expect
			.poll(() =>
				page.evaluate(() =>
					['first-slide', 'select', 'type', 'switch'].every(
						(name) => performance.getEntriesByName(name, 'measure').length === 1,
					),
				),
			)
			.toBe(true);
		const timings = await page.evaluate(() =>
			Object.fromEntries(
				['first-slide', 'select', 'type', 'switch'].map((name) => [
					name,
					performance.getEntriesByName(name, 'measure')[0]?.duration,
				]),
			),
		);
		await info.attach('pptx-opening-timings-ms', {
			body: JSON.stringify({ slides: size, ...timings }),
			contentType: 'application/json',
		});
		expect(errors).toEqual([]);
	});
}

for (const initialSlide of [0, 3]) {
	test(`shows opening slide ${initialSlide + 1} before the editable deck and fills thumbnails`, async ({
		page,
	}) => {
		const title = `Deck slide ${initialSlide + 1} shape 1`;
		// Observe the transient nodes in the page before the editable deck replaces them.
		await page.addInitScript(
			({ preview, editable, title, slideCount }) => {
				const thumbnailSelector = '[data-pptx-preview-thumbnail]';
				let firstFrameScheduled = false;
				const observer = new MutationObserver(() => {
					const opening = document.querySelector(preview);
					if (!opening?.getClientRects().length || document.querySelector(editable)) return;
					if (!opening.firstElementChild?.textContent?.includes(title)) return;
					if (!firstFrameScheduled) {
						firstFrameScheduled = true;
						requestAnimationFrame(() => {
							if (opening.getClientRects().length > 0) {
								document.documentElement.dataset.pptxFirstFrameThumbnailCount = String(
									opening.querySelectorAll(thumbnailSelector).length,
								);
							}
						});
					}
					const thumbnails = opening.querySelectorAll<HTMLElement>(thumbnailSelector);
					if (
						thumbnails.length === slideCount &&
						Array.from(thumbnails).every(
							(thumb, index) =>
								thumb.dataset.pptxPreviewThumbnail === String(index) &&
								thumb.textContent?.includes(`Deck slide ${index + 1} shape 1`),
						)
					) {
						document.documentElement.dataset.pptxObservedOpening = 'true';
						observer.disconnect();
					}
				});
				observer.observe(document, { childList: true, subtree: true });
			},
			{ preview: PREVIEW, editable: CANVAS_ELEMENT, title, slideCount: PPTX_OPEN_SMALL },
		);
		await page.goto(`/?initialSlide=${initialSlide}`);
		await upload(page, PPTX_OPEN_SMALL);
		await expect(
			page.locator(CANVAS_ELEMENT).getByText(title, { exact: true }).first(),
		).toBeVisible();
		await expect(page.locator('html')).toHaveAttribute(
			'data-pptx-first-frame-thumbnail-count',
			'0',
		);
		await expect(page.locator('html')).toHaveAttribute('data-pptx-observed-opening', 'true');
		await expect(page.locator(PREVIEW)).toBeHidden();
	});
}

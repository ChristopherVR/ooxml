/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright spec */
/**
 * A picture's own frame (`a:ln`) and outer shadow (`a:outerShdw`), which a
 * Picture Styles pick writes. PowerPoint centres the line on the frame edge,
 * so half of it lies outside the picture and half over it. Before the fix
 * React painted the photo over the inner half, Vue and Svelte cut the line to
 * its inner half and dropped the shadow, and Angular and vanilla painted
 * neither.
 *
 * Run: bunx playwright test picture-frame-shadow
 */
import { expect, test } from '@playwright/test';

import { resetTabSession, slideStage } from './support/deck';
import { pictureFrameDeck } from './support/picture-frame-deck';

test('a framed picture paints its whole frame and casts its outer shadow', async ({ page }) => {
	await resetTabSession(page);
	await page.goto('/');
	await page.locator('#file-input').setInputFiles({
		name: 'picture-frame.pptx',
		mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
		buffer: await pictureFrameDeck(),
	});
	const img = slideStage(page).locator('[data-element-id] img').first();
	await img.waitFor();
	const result = await img.evaluate((node) => {
		const box = node.closest<HTMLElement>('[data-element-id]')!;
		const clips = (el: Element): boolean => {
			const cs = getComputedStyle(el);
			return cs.overflow !== 'visible' || cs.clipPath !== 'none';
		};
		// The nearest layer that clips the picture's pixels (the crop clip).
		let clip: Element | null = node.parentElement;
		while (clip && clip !== box && !clips(clip)) {
			clip = clip.parentElement;
		}
		const stroke = [...box.querySelectorAll('svg path')].find(
			(p) => p.getAttribute('stroke')?.toUpperCase() === '#FF0000',
		);
		// Every layer from the outline up to the picture box must leave it unclipped,
		// or the outer half of the centred line is cut away.
		let strokeClipped = !stroke;
		for (let el = stroke?.closest('svg')?.parentElement ?? null; el; el = el.parentElement) {
			if (clips(el)) {
				strokeClipped = true;
			}
			if (el === box) {
				break;
			}
		}
		return {
			hasStroke: Boolean(stroke),
			strokeInsideCrop: Boolean(stroke && clip && clip.contains(stroke)),
			strokeClipped,
			// Painted after the picture, so the photo does not cover the inner half.
			strokeAbovePicture: Boolean(
				stroke && node.compareDocumentPosition(stroke) & Node.DOCUMENT_POSITION_FOLLOWING,
			),
			shadow: [box, ...box.querySelectorAll('*')].some((el) =>
				getComputedStyle(el).filter.includes('drop-shadow'),
			),
		};
	});
	expect(result).toEqual({
		hasStroke: true,
		strokeInsideCrop: false,
		strokeClipped: false,
		strokeAbovePicture: true,
		shadow: true,
	});
});

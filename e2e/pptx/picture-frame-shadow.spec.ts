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

for (const animated of [false, true])
	test(`a framed picture paints its whole frame and shadow${animated ? ' during animation' : ''}`, async ({
		page,
	}, testInfo) => {
		await resetTabSession(page);
		await page.goto('/');
		await page.locator('#file-input').setInputFiles({
			name: 'picture-frame.pptx',
			mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
			buffer: await pictureFrameDeck(animated),
		});
		let img = slideStage(page).locator('[data-element-id] img').first();
		await img.waitFor();
		if (animated) {
			const id = await img.evaluate((node) =>
				node.closest('[data-element-id]')!.getAttribute('data-element-id')!,
			);
			const show = page.getByRole('button', { name: /^slide show$/iu });
			if (await show.count()) await show.last().click();
			else
				await page
					.getByRole('button', { name: /present/iu })
					.first()
					.click();
			await page.waitForFunction((id) => {
				const nodes = [...document.querySelectorAll<HTMLElement>('[data-element-id]')].filter(
					(node) => node.dataset.elementId === id,
				);
				for (const node of nodes) node.removeAttribute('data-e2e-animated-frame');
				const box = nodes.sort(
					(a, b) =>
						b.getBoundingClientRect().width * b.getBoundingClientRect().height -
						a.getBoundingClientRect().width * a.getBoundingClientRect().height,
				)[0];
				box?.setAttribute('data-e2e-animated-frame', 'true');
				return box && getComputedStyle(box).visibility === 'hidden';
			}, id);
			await page.keyboard.press('PageDown');
			const box = page.locator('[data-e2e-animated-frame]');
			await expect(box).toBeVisible();
			await expect
				.poll(() => box.evaluate((node) => getComputedStyle(node).animationName))
				.not.toBe('none');
			img = box.locator('img').first();
		}

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
		await page.screenshot({
			path: testInfo.outputPath(animated ? 'animated-frame.png' : 'static-frame.png'),
		});
	});

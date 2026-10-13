/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright spec */
import { writeFile } from 'node:fs/promises';

import { expect, test } from '@playwright/test';
import JSZip from 'jszip';

import { createVisibilityTimingDeck } from '../../src/core/pptx/__tests__/helpers/visibility-timing-deck';
import { savePptxViaBackstage } from './save-pptx';
import { loadDeck } from './support/deck';
import { downloadBytes } from './support/exports';

test.use({ viewport: { width: 1440, height: 900 } });
async function present(page: Parameters<typeof loadDeck>[0]) {
	const show = page.getByRole('button', { name: /^slide show$/iu });
	if (await show.count()) await show.last().click();
	else
		await page
			.getByRole('button', { name: /present/iu })
			.first()
			.click();
	await expect(page.locator('pptx-ui-present-toolbar:visible')).toHaveCount(1);
}
test('native held visibility uses its discrete child clock and supports seek/stop/re-entry/export', async ({
	page,
}, info) => {
	const source = await createVisibilityTimingDeck(1499);
	const input = info.outputPath('long.pptx');
	await writeFile(input, source);
	await loadDeck(page, input);
	const originalZip = await JSZip.loadAsync(source);
	const originalXml = await originalZip.file('ppt/slides/slide1.xml')!.async('string');
	const exported = await downloadBytes(await savePptxViaBackstage(page));
	const savedZip = await JSZip.loadAsync(exported);
	const savedXml = await savedZip.file('ppt/slides/slide1.xml')!.async('string');
	expect(savedXml.match(/<p:timing\b[\s\S]*?<\/p:timing>/u)?.[0]).toBe(
		originalXml.match(/<p:timing\b[\s\S]*?<\/p:timing>/u)?.[0],
	);
	const saved = info.outputPath('saved.pptx');
	await writeFile(saved, exported);
	await loadDeck(page, saved);
	await present(page);
	const target = page.locator('[data-element-id][aria-label="RectDimRGB"]').last();
	await expect(target).toBeHidden();
	await page.keyboard.press('PageDown');
	await expect.poll(() => target.evaluate((node) => node.getAnimations().length)).toBe(1);
	const samples = await target.evaluate(async (node) => {
		const animation = node.getAnimations()[0]!;
		const read = (time: number) => {
			animation.pause();
			animation.currentTime = time;
			const style = getComputedStyle(node);
			return {
				visibility: style.visibility,
				opacity: style.opacity,
				name: style.animationName,
				duration: style.animationDuration,
			};
		};
		const before = read(1000);
		const after = read(1499.5);
		animation.play();
		await animation.finished.catch(() => undefined);
		return { before, after };
	});
	expect(samples.before).toMatchObject({ visibility: 'hidden', opacity: '1', duration: '1.5s' });
	expect(samples.after).toMatchObject({ visibility: 'visible', opacity: '1' });
	expect(samples.before.name).toContain('pptx-tl-visibility');
	await expect(target).toBeVisible();
	await page.keyboard.press('Escape');
	await present(page);
	await expect(target).toBeHidden();
	await page.keyboard.press('PageDown');
	await expect.poll(() => target.evaluate((node) => node.getAnimations().length)).toBe(1);
	// A second next while active seeks this group, without starting another click.
	await page.keyboard.press('PageDown');
	await expect(target).toBeVisible();
	await page.keyboard.press('Escape');
	await present(page);
	await expect(target).toBeHidden();
	await page.keyboard.press('Escape');
});

/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright provides test and expect */
import { test, expect } from '@playwright/test';

import { generateFixture } from './fixtures/generate-arrow-text-regressions';
import { resetTabSession, slideStage } from './support/deck';

test.use({ viewport: { width: 1920, height: 1080 } });

let fixture: string;
test.beforeAll(async () => {
	fixture = await generateFixture();
});

/** Measure painted character positions, independently of a binding's span structure. */
function characterRects(root: Element, text: string): number[][] {
	const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
	const nodes: Text[] = [];
	let combined = '';
	while (walker.nextNode()) {
		nodes.push(walker.currentNode as Text);
		combined += walker.currentNode.textContent;
	}
	const start = combined.indexOf(text);
	if (start < 0) throw new Error(`Missing text: ${text}`);
	const result: number[][] = [];
	let offset = 0;
	for (const node of nodes) {
		for (let i = 0; i < node.length; i += 1) {
			if (offset + i < start || offset + i >= start + text.length) continue;
			if (/\s/u.test(node.data[i]!)) continue;
			const range = document.createRange();
			range.setStart(node, i);
			range.setEnd(node, i + 1);
			const rect = range.getBoundingClientRect();
			result.push([rect.x, rect.y, rect.width, rect.height]);
		}
		offset += node.length;
	}
	return result;
}

test.beforeEach(async ({ page }) => {
	await resetTabSession(page);
	await page.goto('/');
	await page.locator('#file-input').setInputFiles(fixture);
	await page.getByRole('button', { name: 'Go to slide 3', exact: true }).first().waitFor();
	await page.waitForFunction(() => document.fonts.status === 'loaded');
	await page.waitForTimeout(1200);
});

test('letter builds keep character positions and wrapping when the outgoing fade begins', async ({
	page,
}) => {
	await page
		.getByRole('button', { name: /^present$|slide show/iu })
		.first()
		.click();
	await page.locator('[data-anim-id*="::c0-"]').first().waitFor({ state: 'attached' });
	await page.waitForTimeout(3000);
	const texts = [
		'Testing Anatidae Presentation',
		'Image By Ams100272 - Own work, CC BY-SA 4.0, https://commons.wikimedia.org/w/index.php?curid=146841823',
	];
	const before = await Promise.all(
		texts.map((text) => slideStage(page).evaluate(characterRects, text)),
	);
	await page.keyboard.press('ArrowRight');
	const outgoing = page.locator('[data-pptx-transition-layer="outgoing"]');
	await outgoing.waitFor();
	for (const [index, text] of texts.entries()) {
		const after = await outgoing.evaluate(characterRects, text);
		expect(after).toHaveLength(before[index]!.length);
		for (const [at, rect] of after.entries()) {
			for (let axis = 0; axis < 2; axis += 1) {
				expect(Math.abs(rect[axis]! - before[index]![at]![axis]!)).toBeLessThan(0.25);
			}
		}
	}
});

test('open freeforms and straight connectors paint correctly sized arrowheads', async ({
	page,
}) => {
	await page.getByRole('button', { name: 'Go to slide 3', exact: true }).first().click();
	const stage = slideStage(page);
	const freeform = stage
		.locator('[data-element-id]')
		.filter({ has: page.locator('path[marker-end]') })
		.first();
	await expect(freeform.locator('marker')).toHaveCount(1);
	await expect(freeform).not.toHaveCSS('clip-path', /path\(/u);
	const markers = stage.locator('marker');
	await expect(markers).toHaveCount(4);
	const dimensions = await markers.evaluateAll((nodes) =>
		nodes.map((node) => Number(node.getAttribute('markerWidth'))).sort((a, b) => a - b),
	);
	expect(dimensions).toEqual([2, 3, 3, 5]);
});

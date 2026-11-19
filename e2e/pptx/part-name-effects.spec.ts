/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright spec, `test`/`expect` come from @playwright/test */
/**
 * SVG effects are painted through `url(#id)` references built from the element id, and the id
 * carries the slide's part name. Whatever that name holds, the effects must look the same.
 */
import { expect, test, type Page } from '@playwright/test';

import { resetTabSession, slideStage } from './support/deck';
import { deckWithSlidePart } from './support/hostile-part-name-deck';
import { pixelDiff } from './support/pixel-diff';

const ORDINARY_NAME = 'slide1.xml';
const UNUSUAL_NAMES = ['スライド1.xml', 'diapositivé1.xml', 'my slide1.xml', 'my%20slide1.xml'];

/** Share of the stage that may differ between two renders of the same slide. */
const SAME_RENDER_FRACTION = 0.002;
/** Share of the stage the effects must change, so a run without them cannot pass. */
const VISIBLE_EFFECT_FRACTION = 0.005;

function shape(id: number, x: number, color: string, effect: string): string {
	return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Shape ${id}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${x}" y="914400"/><a:ext cx="2743200" cy="2743200"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="${color}"/></a:solidFill>${effect}</p:spPr></p:sp>`;
}

const SOFT_EDGE = '<a:effectLst><a:softEdge rad="254000"/></a:effectLst>';
const BEVEL =
	'<a:scene3d><a:camera prst="orthographicFront"/><a:lightRig rig="threePt" dir="t"/></a:scene3d><a:sp3d><a:bevelT w="190500" h="190500"/></a:sp3d>';

const shapes = (softEdge: string, bevel: string): string =>
	shape(900, 457200, 'C83232', softEdge) + shape(901, 4572000, '3264C8', bevel);

async function stageImage(page: Page, fileName: string, content: string): Promise<Buffer> {
	await resetTabSession(page);
	await page.goto('/');
	await page.locator('#file-input').setInputFiles({
		name: 'part-name-effects.pptx',
		mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
		buffer: await deckWithSlidePart(fileName, content),
	});
	const stage = slideStage(page);
	await expect(stage.locator('[data-element-id]').first()).toBeAttached();
	return stage.screenshot();
}

test('soft edge and bevel look the same under any slide part name', async ({ page }) => {
	const withEffects = shapes(SOFT_EDGE, BEVEL);
	const reference = await stageImage(page, ORDINARY_NAME, withEffects);

	const plain = await stageImage(page, ORDINARY_NAME, shapes('', ''));
	const effectArea = await pixelDiff(page, reference, plain);
	expect(effectArea.diffPixelFraction).toBeGreaterThan(VISIBLE_EFFECT_FRACTION);

	for (const name of UNUSUAL_NAMES) {
		const image = await stageImage(page, name, withEffects);
		const diff = await pixelDiff(page, reference, image);
		expect(diff.diffPixelFraction, name).toBeLessThan(SAME_RENDER_FRACTION);
	}
});

/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright spec */
import { expect, test } from '@playwright/test';
import type { Page } from '@playwright/test';

import { TWO_THEMES } from './fixtures/generate-two-themes-fixture';
import { savePptxViaBackstage } from './save-pptx';
import {
	fixture,
	inspector,
	loadDeck,
	openRibbonTab,
	slideElements,
	thumbnail,
} from './support/deck';
import { downloadBytes } from './support/exports';
import { slidePosition } from './support/keyboard';
import { readZipPartText } from './support/pptx-xml';
import { chooseSelectValue } from './support/select-control';

/**
 * Presentation properties > Theme: pick another theme part already in the package and apply it to
 * every master. It applies in place, as one undo step: the slide on screen stays, the slides drawn
 * from the master that moved take the new theme's colours, Undo puts the old theme back, and a save
 * writes the masters' theme relationships as switched.
 *
 * `two-themes.pptx`: slides 1 and 2 on master 1 ("Tide", accent 1 `1F6FB2`), slide 3 on master 2
 * ("Ember", accent 1 `C0392B`); every slide has an `accent1` rectangle, the largest filled shape on
 * it. Framework-neutral: it reads the inspector's theme select and buttons by their accessible
 * names, the slide counter, the shapes' painted background and the saved file.
 */
test.use({ viewport: { width: 1440, height: 900 } });

const { theme1, theme2 } = TWO_THEMES;

const rgb = (hex: string) =>
	`rgb(${[0, 2, 4].map((at) => Number.parseInt(hex.slice(at, at + 2), 16)).join(', ')})`;

/** The painted background of the largest filled shape on the slide: its `accent1` rectangle. */
function accentFill(page: Page): Promise<string> {
	return slideElements(page).evaluateAll((nodes) => {
		let best = { area: 0, color: 'none' };
		for (const node of nodes) {
			const color = getComputedStyle(node).backgroundColor;
			const box = node.getBoundingClientRect();
			const area = box.width * box.height;
			if (color !== 'rgba(0, 0, 0, 0)' && color !== 'transparent' && area > best.area) {
				best = { area, color };
			}
		}
		return best.color;
	});
}

/** The theme part each slide master's relationship points at, as File > Save writes it. */
async function savedMasterThemes(page: Page): Promise<string[]> {
	const bytes = await downloadBytes(await savePptxViaBackstage(page));
	const targets: string[] = [];
	for (const master of [1, 2]) {
		const rels = await readZipPartText(
			bytes,
			`ppt/slideMasters/_rels/slideMaster${master}.xml.rels`,
		);
		targets.push(/Target="[^"]*\/(theme\d+\.xml)"/u.exec(rels)?.[1] ?? 'missing');
	}
	return targets;
}

test('a theme applied by path to every master re-themes in place, and Undo takes it back', async ({
	page,
}) => {
	await loadDeck(page, fixture('two-themes.pptx'));
	await thumbnail(page, 2).click();
	await expect.poll(() => slidePosition(page)).toMatch(/^Slide 2 of 3/u);
	await expect.poll(() => accentFill(page)).toBe(rgb(theme1.colors.accent1));

	// Design > Slide Size shows the inspector's Presentation properties (nothing selected).
	await openRibbonTab(page, 'Design');
	await page.getByRole('button', { name: 'Slide Size', exact: true }).first().click();
	const select = inspector(page)
		.locator('[aria-label="Theme"]')
		.filter({ has: page.locator(`option[value="${theme2.path}"]`) })
		.first();
	await chooseSelectValue(page, select, theme2.path);
	await inspector(page).getByRole('button', { name: 'Apply All Masters', exact: true }).click();

	// React, Vue and Svelte reloaded the deck (back on slide 1, nothing to undo); Angular and
	// vanilla re-pointed the masters but left the slides in theme 1's colours.
	await expect.poll(() => accentFill(page)).toBe(rgb(theme2.colors.accent1));
	expect(await slidePosition(page)).toMatch(/^Slide 2 of 3/u);
	const undo = page.getByRole('button', { name: 'Undo', exact: true }).first();
	await expect(undo).toBeEnabled();
	expect(await savedMasterThemes(page)).toStrictEqual(['theme2.xml', 'theme2.xml']);

	await undo.click();
	await expect.poll(() => accentFill(page)).toBe(rgb(theme1.colors.accent1));
	expect(await slidePosition(page)).toMatch(/^Slide 2 of 3/u);
	expect(await savedMasterThemes(page)).toStrictEqual(['theme1.xml', 'theme2.xml']);
});

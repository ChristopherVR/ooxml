/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright spec */
import { test, expect } from '@playwright/test';

import { loadDeck } from './support/deck';

/** The Insert commands checked, with the group that holds each. */
const INSERT_GROUP_OF: Readonly<Record<string, string>> = {
	'Text Box': 'insert.text',
	Pictures: 'insert.images',
	Video: 'insert.media',
	Table: 'insert.tables',
	SmartArt: 'insert.illustrations',
	Equation: 'insert.symbols',
};

/** Design groups also collapse right to left when the window is narrow, as in Office. */
const DESIGN_GROUP_OF: Readonly<Record<string, string>> = {
	'Browse Themes': 'design.themes',
	'Edit Theme': 'design.themes',
	'Slide Size': 'design.customize',
	'Format Background': 'design.customize',
};

/**
 * The command's control box (a split button's face and chevron together) and whether its
 * group is scaled down (`data-scale`), crossing the shadow roots some commands render in.
 */
function controlSize(button: Element): { height: number; scaled: boolean } {
	const control = button.closest('.menu.split') ?? button;
	let node: Element | null = button;
	while (node && !node.hasAttribute('data-ribbon-group')) {
		node = node.parentElement ?? ((node.getRootNode() as ShadowRoot).host || null);
	}
	return {
		height: control.getBoundingClientRect().height,
		scaled: node?.hasAttribute('data-scale') ?? false,
	};
}

// A narrow window scales the Insert groups as Office does (large, then small with labels, then
// icons only) before any becomes a popup, so an Insert command's size follows its group's scale;
// a popup draws its group large.
for (const width of [1000, 1280]) {
	test.describe(`compact ribbon at ${width}px`, () => {
		test.use({ viewport: { width, height: 800 } });

		for (const [tab, commands, large] of [
			['Design', ['Browse Themes', 'Edit Theme', 'Slide Size', 'Format Background'], true],
			['Insert', Object.keys(INSERT_GROUP_OF), true],
			['Transitions', ['Apply to All'], false],
		] as const) {
			test(`${tab} commands keep their Office size and alignment`, async ({ page }, testInfo) => {
				await loadDeck(page);
				const toolbar = page.getByRole('toolbar', { name: 'Presentation toolbar' });
				await toolbar.getByRole('tab', { name: tab, exact: true }).click();
				for (const name of commands) {
					// Near 1000px a group can become a popup (font metrics decide), as in Office: open it.
					const group = INSERT_GROUP_OF[name] ?? DESIGN_GROUP_OF[name];
					const collapsed = toolbar.locator(`[data-ribbon-group="${group}"][data-collapsed]`);
					const popup = group !== undefined && (await collapsed.count()) > 0;
					if (popup) {
						await collapsed.locator('[data-pptx-chrome="ribbon-collapse"]').click();
					}
					const command = toolbar.getByRole('button', { name, exact: true });
					await expect(command).toBeVisible();
					const { height, scaled } = await command.evaluate(controlSize);
					if (popup) {
						await page.keyboard.press('Escape');
					}
					// Large commands are Office's 66px glyph-over-caption tiles; the rest stay 24-36px rows.
					const tile = large && !scaled;
					expect(height, `${name} must not stretch to the ribbon height`).toBeLessThanOrEqual(
						tile ? 66 : 36,
					);
					expect(height, `${name} height`).toBeGreaterThanOrEqual(tile ? 66 : 24);
				}
				if (tab === 'Transitions') {
					// Preview is a large (icon above label) command, as in PowerPoint.
					const preview = await toolbar
						.getByRole('button', { name: 'Preview', exact: true })
						.boundingBox();
					expect(preview!.height).toBeGreaterThanOrEqual(44);
				}
				const overflow = await page.evaluate(
					() => document.documentElement.scrollWidth - document.documentElement.clientWidth,
				);
				expect(
					overflow,
					'ribbon must scroll internally without widening the page',
				).toBeLessThanOrEqual(1);
				await page.screenshot({ path: testInfo.outputPath(`${tab}.png`) });
			});
		}
	});
}

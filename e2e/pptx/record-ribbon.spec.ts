/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright spec */
import { expect, test } from '@playwright/test';

import { loadDeck } from './support/deck';
import { collectRibbonInventory } from './support/ribbon-controls';

test.use({ viewport: { width: 1440, height: 900 } });
test('Record ribbon retains its six commands and four disabled placeholders', async ({
	page,
}, testInfo) => {
	await loadDeck(page);
	const inventory = await collectRibbonInventory(page, ['Record']);
	expect(inventory[0].controls).toHaveLength(6);
	expect(inventory[0].controls.filter((control) => control.disabled)).toHaveLength(4);
	const toolbar = page.getByRole('toolbar', { name: 'Presentation toolbar' });
	await expect(toolbar.locator('pptx-ui-ribbon-group:visible')).toHaveCount(4);
	// Every binding renders the tab through the shared section.
	await expect(
		toolbar.locator('pptx-ui-ribbon-section > pptx-ui-ribbon-group:visible'),
	).toHaveCount(4);
	// The section fills the ribbon row as a group does, so the three-line From Current Slide
	// label clears the group caption.
	const fill = await toolbar.locator('pptx-ui-ribbon-section:visible').evaluate((section) => {
		// The row is the nearest ancestor with a box (a component host can be display: contents).
		let row = section.parentElement!;
		while (getComputedStyle(row).display === 'contents') {
			row = row.parentElement!;
		}
		const style = getComputedStyle(row);
		const content =
			row.clientHeight - parseFloat(style.paddingTop) - parseFloat(style.paddingBottom);
		return { section: section.getBoundingClientRect().height, content };
	});
	expect(fill.section).toBeCloseTo(fill.content, 0);
	await expect(toolbar.locator('pptx-ui-ribbon-command:visible')).toHaveCount(6);
	const beginning = toolbar.getByRole('button', { name: 'From Beginning', exact: true });
	const current = toolbar.getByRole('button', { name: 'From Current Slide', exact: true });
	await beginning.focus();
	await page.keyboard.press('Tab');
	await expect(current).toBeFocused();
	await expect(toolbar.getByRole('button', { name: 'Cameo', exact: true })).toBeDisabled();
	await expect(toolbar.getByRole('button', { name: 'Learn More', exact: true })).toBeDisabled();
	const path = testInfo.outputPath('record-ribbon.png');
	await toolbar.screenshot({ path });
	await testInfo.attach('record-ribbon', { path, contentType: 'image/png' });
});

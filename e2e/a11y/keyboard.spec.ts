/**
 * The keyboard contract in the suite, for Word, Excel, PowerPoint, Visio and Teams: Tab goes round
 * the page without a trap, F6 reaches every region, dialogs contain focus and give it back, and
 * saving leaves focus where it was. Tab order only exists in a real browser, so it is proved here.
 */
import { SUITE, type SuiteProduct } from './support/contracts';
import { dialogRun, focusStop, regionCycle, tabCycle } from './support/keyboard';
import { gate } from './support/known';
import { expect, openSurface, test } from './support/suite';

test.use({ viewport: { width: 1440, height: 900 } });

const unsaved = (page: import('@playwright/test').Page) =>
	page.evaluate(() =>
		Boolean(
			document
				.querySelector('#tabs [role="tab"][aria-selected="true"]')
				?.textContent?.includes('•'),
		),
	);

for (const [product, contract] of Object.entries(SUITE) as [
	SuiteProduct,
	(typeof SUITE)[SuiteProduct],
][]) {
	test.describe(`${contract.title} keyboard contract`, () => {
		test.beforeEach(async ({ page }) => {
			await openSurface(page, product);
		});

		test('Tab goes round the page without a trap and through every region', async ({ page }) => {
			await gate(`keyboard/${product}/tab`, async () => {
				const cycle = await tabCycle(page, contract);
				const trail = cycle.trail.slice(-12).join('\n');
				expect(cycle.trap, `Tab stuck on ${cycle.trap}\n${trail}`).toBeUndefined();
				expect(cycle.closed, `Tab never came back to the first stop\n${trail}`).toBe(true);
				expect([...cycle.regions].sort()).toEqual(
					expect.arrayContaining([...contract.tabRegions].sort()),
				);
			});
		});

		test('F6 and Shift+F6 reach every region and come back to the document', async ({ page }) => {
			await gate(`keyboard/${product}/f6`, async () => {
				for (const key of ['F6', 'Shift+F6'] as const) {
					const cycle = await regionCycle(page, contract, key);
					const trail = cycle.trail.join(' | ');
					expect(
						cycle.returned,
						`${key} never came back to the ${contract.document}: ${trail}`,
					).toBe(true);
					expect(cycle.order, `${key}: ${trail}`).toEqual(
						expect.arrayContaining([...contract.f6Regions]),
					);
				}
			});
		});

		test('a dialog keeps focus inside, closes on Escape and gives focus back', async ({ page }) => {
			await gate(`keyboard/${product}/dialog`, async () => {
				const run = await dialogRun(page, contract, contract.dialog);
				expect(run.focusMovedIn, 'opening the dialog moves focus into it').toBe(true);
				expect(run.escapes, 'Tab and Shift+Tab stay inside the dialog').toEqual([]);
				expect(run.closedByEscape, 'Escape closes the dialog').toBe(true);
				if (contract.dialog.returnsTo === 'opener')
					expect(run.after.id, `focus returns to ${run.before.label}, not ${run.after.label}`).toBe(
						run.before.id,
					);
				else
					expect(run.after.region, `focus returns to the ${contract.document}`).toBe(
						contract.document,
					);
			});
		});

		if (contract.edit) {
			const edit = contract.edit;
			test('saving an edit keeps keyboard focus where it was', async ({ page }) => {
				await gate(`keyboard/${product}/save`, async () => {
					await edit(page);
					await expect.poll(() => unsaved(page), 'the edit marks the document unsaved').toBe(true);
					const before = await focusStop(page, contract);
					expect(before.region).toBe(contract.document);
					// Straight away: a save asked for while the edit is still being written waits for it.
					await page.keyboard.press('Control+s');
					await expect(
						page.locator('footer.statusbar #status'),
						'Ctrl+S saves the document',
					).toHaveText('Saved to this device');
					await expect.poll(() => unsaved(page), 'the tab shows the document saved').toBe(false);
					await expect(page.locator('.document-host[aria-busy]'), 'the save is done').toHaveCount(
						0,
					);
					const after = await focusStop(page, contract);
					expect(after.key, `focus stays on ${before.label}, not ${after.label}`).toBe(before.key);
				});
			});
		}
	});
}

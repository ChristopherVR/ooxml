/**
 * The PowerPoint keyboard contract in every binding's demo (one Playwright project each): Tab goes
 * round the page without a trap, F6 reaches every region, the shortcut reference dialog contains
 * focus and gives it back, and an edit keeps focus on the slide.
 */
import { POWERPOINT_REGIONS, powerPoint } from './support/contracts';
import { powerPointDeck } from './support/fixtures';
import { dialogRun, focusStop, regionCycle, tabCycle } from './support/keyboard';
import { gate } from './support/known';
import { expect, test } from './support/suite';

test.use({ viewport: { width: 1440, height: 900 } });

const contract = powerPoint(POWERPOINT_REGIONS, [
	'ribbon',
	'thumbnails',
	'slide',
	'format pane',
	'notes',
	'status',
]);

test.describe('PowerPoint keyboard contract', () => {
	test.beforeEach(async ({ page }) => {
		await page.goto('/');
		await page.locator('#file-input').setInputFiles(powerPointDeck());
		await page.locator('[data-pptx-element="true"]').first().waitFor({ timeout: 60_000 });
		await page.evaluate(() => document.fonts.ready);
		await page.waitForTimeout(500);
	});

	const key = (check: string) => `keyboard/${test.info().project.name}/${check}`;

	test('Tab goes round the page without a trap and through every region', async ({ page }) => {
		await gate(key('tab'), async () => {
			const cycle = await tabCycle(page, contract);
			const trail = cycle.trail.slice(-12).join('\n');
			expect(cycle.trap, `Tab stuck on ${cycle.trap}\n${trail}`).toBeUndefined();
			expect(cycle.closed, `Tab never came back to the first stop\n${trail}`).toBe(true);
			expect([...cycle.regions]).toEqual(expect.arrayContaining([...contract.tabRegions]));
		});
	});

	test('F6 and Shift+F6 reach every region and come back to the slide', async ({ page }) => {
		await gate(key('f6'), async () => {
			for (const press of ['F6', 'Shift+F6'] as const) {
				const cycle = await regionCycle(page, contract, press);
				const trail = cycle.trail.join(' | ');
				expect(cycle.returned, `${press} never came back to the slide: ${trail}`).toBe(true);
				expect(cycle.order, `${press}: ${trail}`).toEqual(
					expect.arrayContaining([...contract.f6Regions]),
				);
			}
		});
	});

	test('a dialog keeps focus inside, closes on Escape and gives focus back', async ({ page }) => {
		await gate(key('dialog'), async () => {
			const run = await dialogRun(page, contract, contract.dialog);
			expect(run.focusMovedIn, 'opening the dialog moves focus into it').toBe(true);
			expect(run.escapes, 'Tab and Shift+Tab stay inside the dialog').toEqual([]);
			expect(run.closedByEscape, 'Escape closes the dialog').toBe(true);
			expect(run.after.id, `focus returns to ${run.before.label}, not ${run.after.label}`).toBe(
				run.before.id,
			);
		});
	});

	// The backstage is a modal page holding controls inside shadow roots (its search box).
	test('the File backstage keeps focus inside and Escape returns to the File tab', async ({
		page,
	}) => {
		await gate(key('backstage'), async () => {
			const run = await dialogRun(page, contract, {
				opener: (p) => p.getByRole('tab', { name: 'File', exact: true }),
				returnsTo: 'opener',
			});
			expect(run.focusMovedIn, 'the File tab moves focus into the backstage').toBe(true);
			expect(run.escapes, 'Tab and Shift+Tab stay inside the backstage').toEqual([]);
			expect(run.closedByEscape, 'Escape closes the backstage').toBe(true);
			expect(run.after.id, `focus returns to ${run.before.label}, not ${run.after.label}`).toBe(
				run.before.id,
			);
		});
	});

	test('an edit from the keyboard keeps focus on the slide', async ({ page }) => {
		await gate(key('edit'), async () => {
			const before = await page.locator('[data-pptx-element="true"]').first().boundingBox();
			await contract.edit(page);
			await expect
				.poll(
					async () => (await page.locator('[data-pptx-element="true"]').first().boundingBox())?.x,
				)
				.not.toBe(before?.x);
			await page.waitForTimeout(300);
			expect((await focusStop(page, contract)).region).toBe('slide');
		});
	});
});

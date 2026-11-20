/**
 * The PowerPoint keyboard contract in every binding's demo (one Playwright project each): Tab goes
 * round the page without a trap, F6 reaches every region, the shortcut reference dialog contains
 * focus and gives it back, and an edit keeps focus on the slide.
 */
import { POWERPOINT_REGIONS, powerPoint } from './support/contracts';
import { powerPointDeck, powerPointTextDeck } from './support/fixtures';
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

test.describe('PowerPoint text and menus from the keyboard', () => {
	const key = (check: string) => `keyboard/${test.info().project.name}/${check}`;

	test('Tab in a text box types a tab and keeps the edit, as in PowerPoint', async ({ page }) => {
		await gate(key('text-tab'), async () => {
			await page.goto('/');
			await page.locator('#file-input').setInputFiles(powerPointTextDeck());
			await page.locator('[data-pptx-element="true"]').first().waitFor({ timeout: 60_000 });
			// Enter goes into the slide and selects the first shape; Enter again edits its text.
			await page.locator('[data-pptx-viewport]').first().focus();
			await page.keyboard.press('Enter');
			await page.keyboard.press('Enter');
			const editor = page.locator('[data-inline-editor]');
			await expect(editor).toBeFocused();
			await page.keyboard.press('End');
			await page.keyboard.press('Tab');
			await expect(editor, 'Tab keeps the caret in the text box').toBeFocused();
			await expect
				.poll(() => editor.evaluate((node) => node.textContent?.includes('\t')))
				.toBe(true);
			await page.keyboard.press('Shift+Tab');
			await expect(editor, 'Shift+Tab keeps the caret in the text box').toBeFocused();
		});
	});

	for (const label of ['More actions', 'Presentation options']) {
		test(`Escape closes ${label} and gives focus back to its button`, async ({ page }) => {
			await gate(key(`menu-escape/${label}`), async () => {
				await page.goto('/');
				await page.locator('#file-input').setInputFiles(powerPointDeck());
				await page.locator('[data-pptx-element="true"]').first().waitFor({ timeout: 60_000 });
				const trigger = page.getByRole('button', { name: label, exact: true }).first();
				const buttons = page.locator('button:visible');
				const closed = await buttons.count();
				await trigger.click();
				await expect.poll(() => buttons.count(), 'the menu opens').toBeGreaterThan(closed);
				await page.keyboard.press('Escape');
				await expect.poll(() => buttons.count(), 'Escape closes the menu').toBe(closed);
				await expect(trigger, 'focus returns to the button that opened it').toBeFocused();
			});
		});
	}
});

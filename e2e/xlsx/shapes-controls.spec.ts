import { expect, test, type Locator, type Page } from '@playwright/test';
import { openFixture } from './fixture-files';
import { FRAMEWORKS, editor, formulaBarValue, goToCell, pageErrors } from './helpers';

// excel-shapes.xlsx (generate-excel-shapes.ps1): sheet "Shapes" with the list A2:A4 (Red, Green,
// Blue), preset shapes, a text box, a connector and a group, and form controls linked to C1 (check
// box, TRUE), C2 (two option buttons), C3 (drop-down, 2), C4 (list box), C5 (spin button, 3) and C6
// (scroll bar, 40).
const objects = (page: Page, selector: string): Locator =>
	editor(page).locator(`[part="grid"] ${selector}`);
const control = (page: Page, type: string): Locator =>
	objects(page, `.xg-formcontrol[data-control-type="${type}"]`).first();

async function cellText(page: Page, ref: string): Promise<string> {
	await goToCell(page, ref);
	return formulaBarValue(page);
}

for (const framework of FRAMEWORKS)
	test(`${framework}: shapes are drawn and form controls write their linked cell`, async ({
		page,
	}) => {
		const errors = pageErrors(page);
		await openFixture(page, framework, 'excel-shapes.xlsx');

		// Shapes: the rectangle, oval, arrow, text box, connector and group each paint SVG.
		const shapes = objects(page, '.xg-shape');
		await expect(shapes.first()).toBeVisible();
		const painted = await shapes.evaluateAll((nodes) =>
			nodes.map((node) => ({
				label: node.getAttribute('aria-label'),
				svg: node.querySelector('svg') !== null,
			})),
		);
		expect(painted.every((s) => s.svg)).toBe(true);
		expect(painted.map((s) => s.label)).toEqual(
			expect.arrayContaining(['Box', 'Oval', 'Arrow', 'Note', 'Line', 'Pair']),
		);
		await expect(shapes.filter({ hasText: 'Hello' }).first()).toBeVisible();

		// Check box: pressing it flips the linked cell.
		expect(await cellText(page, 'C1')).toBe('TRUE');
		await control(page, 'checkbox').locator('[data-control-part="toggle"]').first().click();
		await expect.poll(() => cellText(page, 'C1')).toBe('FALSE');
		await control(page, 'checkbox').locator('[data-control-part="toggle"]').first().click();
		await expect.poll(() => cellText(page, 'C1')).toBe('TRUE');

		// Drop-down: open the list and choose the third item.
		expect(await cellText(page, 'C3')).toBe('2');
		await control(page, 'dropdown').locator('[data-control-part="drop"]').first().click();
		await control(page, 'dropdown')
			.locator('[data-control-part="item"][data-control-index="3"]')
			.click();
		await expect.poll(() => cellText(page, 'C3')).toBe('3');

		// Spin button: one step up from 3.
		await control(page, 'spin').locator('[data-control-part="step-up"]').first().click();
		await expect.poll(() => cellText(page, 'C5')).toBe('4');
		expect(errors).toEqual([]);
	});

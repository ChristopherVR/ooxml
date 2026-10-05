import { expect, test } from '@playwright/test';
import {
	FRAMEWORKS,
	SAMPLE_SHEETS,
	editor,
	editorProperty,
	grid,
	newWorkbook,
	openLanding,
	openSample,
	pageErrors,
	part,
	typeInActiveCell,
} from './helpers';

// One editor, six thin bindings: every adapter must mount the same element and its parts.
for (const framework of FRAMEWORKS) {
	test(`${framework}: mounts <xlsx-editor> and opens the sample`, async ({ page }) => {
		const errors = pageErrors(page);
		await openSample(page, framework);
		await expect(editor(page)).toHaveCount(1);
		await expect(page.locator('#build-stamp')).toHaveText(new RegExp(framework, 'u'));
		for (const name of ['ribbon', 'name-box', 'formula-bar', 'grid', 'sheet-tabs', 'status-bar'])
			await expect(part(page, name)).toHaveCount(1);
		for (const name of SAMPLE_SHEETS) await expect(part(page, 'sheet-tabs')).toContainText(name);
		expect(errors).toEqual([]);
	});

	test(`${framework}: a new workbook accepts typed input`, async ({ page }) => {
		await newWorkbook(page, framework);
		await typeInActiveCell(page, `${framework} works`);
		await expect(grid(page)).toContainText(`${framework} works`);
	});

	test(`${framework}: edits recalculate and the binding hears selection and dirty`, async ({
		page,
	}) => {
		const errors = pageErrors(page);
		await openSample(page, framework);
		const stamp = page.locator('#build-stamp');
		await editor(page).evaluate((node) =>
			(node as unknown as { select(ref: string): void }).select('C11'),
		);
		await expect(stamp).toHaveAttribute('data-selection', 'C11');
		await typeInActiveCell(page, '=2*21');
		await expect(grid(page)).toContainText('42');
		await expect(stamp).toHaveAttribute('data-dirty', 'true');
		await expect(stamp).toHaveAttribute('data-selection', 'C12');
		expect(errors).toEqual([]);
	});
}

test('the chosen locale reaches the editor mounted through an adapter', async ({ page }) => {
	await openLanding(page, 'react');
	await page.locator('#locale-select').selectOption('fr');
	await page.locator('#sample').click();
	await expect(grid(page)).toBeVisible();
	expect(await editorProperty<string>(page, 'locale')).toBe('fr');
});

test('switching the locale relabels the ribbon live', async ({ page }) => {
	await openSample(page);
	await expect(part(page, 'ribbon').getByRole('tab', { name: 'Home' })).toBeVisible();
	await editor(page).evaluate((node) => node.setAttribute('locale', 'de'));
	await expect(part(page, 'ribbon').getByRole('tab', { name: 'Start' })).toBeVisible();
	await expect(part(page, 'status-bar')).toContainText('Bereit');
});

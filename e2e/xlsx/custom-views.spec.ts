import { expect, test, type Page } from '@playwright/test';
import { openFixture } from './fixture-files';
import { clickRibbonControl, editor, FRAMEWORKS, pageErrors, ribbon } from './helpers';

// excel-custom-views.xlsx (generate-excel-custom-views.ps1): sheet "Data" holds R<row>C<col> text in
// A1:D6. The view "Compact" hides row 3 and column C, filters column B to R2C2, zooms to 85% and
// hides gridlines. "Full" was saved without hidden rows, columns and filter settings, so (as in
// Excel) showing it leaves the hidden state alone; a view added with them restores it.
const dialog = (page: Page) => editor(page).locator('[data-dialog="custom-views"]');
const button = (page: Page, label: string) =>
	dialog(page).locator('button.xve-btn', { hasText: label });
const view = (page: Page, name: string) => dialog(page).getByRole('option', { name, exact: true });
const names = async (page: Page) =>
	(await dialog(page).getByRole('option').allTextContents()).sort();

/** The "R<row>C<col>" texts of the cells the grid shows (hidden rows and columns are not painted). */
const shown = (page: Page): Promise<string[]> =>
	editor(page).evaluate((node) =>
		[...node.shadowRoot!.querySelectorAll<HTMLElement>('.xg-c:not([hidden])')]
			.map((cell) => cell.textContent ?? '')
			.filter((text) => /^R\dC\d$/.test(text))
			.sort(),
	);
const ALL = [1, 2, 3, 4, 5, 6].flatMap((r) => [1, 2, 3, 4].map((c) => `R${r}C${c}`));
const COMPACT = ['R1C1', 'R1C2', 'R1C4', 'R2C1', 'R2C2', 'R2C4'];

async function openCustomViews(page: Page) {
	await ribbon(page).getByRole('tab', { name: 'View', exact: true }).click();
	await clickRibbonControl(page, ribbon(page).locator('[data-command="view.custom-views"]'));
	await expect(dialog(page)).toBeVisible();
}

const show = async (page: Page, name: string) => {
	await openCustomViews(page);
	await view(page, name).click();
	await button(page, 'Show').click();
	await expect(dialog(page)).toBeHidden();
};

const addView = async (page: Page, name: string) => {
	await openCustomViews(page);
	await button(page, 'Add...').click();
	const add = editor(page).locator('[data-dialog="custom-view-add"]');
	await add.getByLabel('Name:').fill(name);
	await add.getByRole('button', { name: 'OK', exact: true }).click();
};

for (const framework of FRAMEWORKS)
	test(`${framework}: Custom Views add, show and delete restore hidden rows`, async ({ page }) => {
		const errors = pageErrors(page);
		await openFixture(page, framework, 'excel-custom-views.xlsx');
		await expect.poll(() => shown(page)).toEqual(ALL);

		await openCustomViews(page);
		await expect.poll(() => names(page)).toEqual(['Compact', 'Full']);
		await button(page, 'Close').click();

		// Add a view of the untouched sheet (hidden state included), then show Compact.
		await addView(page, 'Mine');
		await expect.poll(() => names(page)).toEqual(['Compact', 'Full', 'Mine']);
		await button(page, 'Close').click();
		await show(page, 'Compact');
		await expect.poll(() => shown(page)).toEqual(COMPACT);

		// Full carries no hidden state: the rows and columns stay hidden.
		await show(page, 'Full');
		await expect.poll(() => shown(page)).toEqual(COMPACT);

		// Mine brings the hidden rows, columns and filter back.
		await show(page, 'Mine');
		await expect.poll(() => shown(page)).toEqual(ALL);

		// Delete removes it from the list.
		await openCustomViews(page);
		await view(page, 'Mine').click();
		await button(page, 'Delete').click();
		await expect.poll(() => names(page)).toEqual(['Compact', 'Full']);
		await button(page, 'Close').click();
		expect(errors).toEqual([]);
	});

import { expect, test } from '@playwright/test';
import { clickRibbonControl, editor, newWorkbook, ribbon, statusBar } from './helpers';

const PNG = Buffer.from(
	'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+jq2sAAAAASUVORK5CYII=',
	'base64',
);

test('edits print settings, worksheet background, note and header from the ribbon', async ({
	page,
}) => {
	await page.setViewportSize({ width: 1640, height: 900 });
	await newWorkbook(page);
	const workbook = editor(page);
	const band = ribbon(page);
	const control = (id: string) => band.locator(`[data-command="${id}"]`);

	await band.getByRole('tab', { name: 'Page Layout', exact: true }).click();
	await control('page.fit-width').selectOption('2');
	await control('page.fit-height').selectOption('1');
	await expect(control('page.fit-width')).toHaveValue('2');
	await expect(control('page.fit-height')).toHaveValue('1');
	await clickRibbonControl(page, control('page.print-titles'));
	const setup = workbook.locator('[data-dialog="page-setup"]');
	await expect(setup.getByRole('tab', { name: 'Sheet' })).toHaveAttribute('aria-selected', 'true');
	await setup.getByLabel('Rows to repeat at top').fill('$1:$2');
	await setup.getByRole('button', { name: 'OK', exact: true }).click();
	await expect(setup).toBeHidden();

	const choosing = page.waitForEvent('filechooser');
	await clickRibbonControl(page, control('page.background'));
	await (await choosing).setFiles({ name: 'background.png', mimeType: 'image/png', buffer: PNG });
	const surface = workbook.locator('.xg-view');
	await expect(surface).toHaveAttribute('data-sheet-background', '');
	await expect(surface.locator('.xg-q-main')).toHaveCSS('background-image', /data:image\/png/);
	await clickRibbonControl(page, control('page.background-delete'));
	await expect(surface).not.toHaveAttribute('data-sheet-background', '');

	await band.getByRole('tab', { name: 'Review', exact: true }).click();
	await clickRibbonControl(page, band.getByRole('button', { name: 'Notes', exact: true }));
	await workbook.getByRole('menuitem', { name: 'New Note', exact: true }).click();
	const note = workbook.locator('[data-dialog="note"]');
	await note.getByLabel('Note', { exact: true }).fill('Review the assumptions');
	await note.getByRole('button', { name: 'OK', exact: true }).click();
	await expect(note).toBeHidden();
	await band.getByRole('tab', { name: 'Insert', exact: true }).click();
	await clickRibbonControl(page, control('insert.header-footer'));
	await expect(statusBar(page).locator('[data-id="pageLayout"]')).toHaveAttribute(
		'aria-pressed',
		'true',
	);
	const input = workbook.locator('.xpl-header input[data-section="center"]');
	await expect(input).toBeFocused();
	await input.fill('Forecast &P');
	await input.press('Enter');
	await expect(workbook.locator('.xpl-header').first()).toContainText('Forecast 1');
});

test('inserts worksheet shapes and text and groups drawings from the Selection Pane', async ({
	page,
}) => {
	await page.setViewportSize({ width: 1700, height: 900 });
	await newWorkbook(page);
	const workbook = editor(page);
	const band = ribbon(page);
	await band.getByRole('tab', { name: 'Insert', exact: true }).click();
	const shapes = band.getByRole('button', { name: 'Shapes', exact: true });
	await clickRibbonControl(page, shapes);
	await band.getByRole('button', { name: 'Rectangle', exact: true }).click();
	await clickRibbonControl(page, shapes);
	await band.getByRole('button', { name: 'Oval', exact: true }).click();
	// Shape Format repeats Text Box; use the Insert tab's Text group.
	await clickRibbonControl(
		page,
		band
			.getByRole('group', { name: 'Text', exact: true })
			.locator('button[data-command="insert.text-box"]'),
	);
	const dialog = workbook.locator('[data-dialog="text-box"]');
	await dialog.getByLabel('Text', { exact: true }).fill('Quarterly forecast');
	await dialog.getByRole('button', { name: 'OK', exact: true }).click();
	await expect(workbook.locator('.xg-q-main .xg-obj')).toHaveCount(3);

	await band.getByRole('tab', { name: 'Page Layout', exact: true }).click();
	// Shape Format shares the Arrange group; use the Page Layout panel's copy.
	const pageLayout = band.locator('#xve-panel-page-layout');
	await clickRibbonControl(
		page,
		pageLayout.locator('button[data-command="arrange.selection-pane"]'),
	);
	const pane = workbook.getByRole('complementary', { name: 'Selection Pane' });
	await pane.getByRole('button', { name: 'Rectangle', exact: true }).click();
	await pane.getByRole('button', { name: 'Oval', exact: true }).click({ modifiers: ['Control'] });
	await clickRibbonControl(page, pageLayout.getByRole('button', { name: 'Group', exact: true }));
	await page.getByRole('menuitem', { name: 'Group', exact: true }).click();
	await expect(workbook.locator('.xg-q-main .xg-obj')).toHaveCount(2);
	await expect(pane.getByRole('button', { name: 'Group', exact: true })).toBeVisible();
	await pane.getByRole('button', { name: 'Hide Drawing: Group', exact: true }).click();
	await expect(workbook.locator('.xg-q-main .xg-obj')).toHaveCount(1);
	await pane.getByRole('button', { name: 'Show Drawing: Group', exact: true }).click();
	await expect(workbook.locator('.xg-q-main .xg-obj')).toHaveCount(2);
});

import { expect, test } from '@playwright/test';
import { goToCell, grid, newWorkbook, pageErrors, typeInActiveCell } from './helpers';

test('REGEXEXTRACT spills all matches and refreshes the painted grid after an edit', async ({
	page,
}) => {
	const errors = pageErrors(page);
	await newWorkbook(page);
	await typeInActiveCell(page, 'order 001, 22, 333');
	await goToCell(page, 'B1');
	await typeInActiveCell(page, '=REGEXEXTRACT(A1,"[0-9]+",1)');
	const visibleCells = grid(page).locator('.xg-c:not([hidden])');
	await expect(visibleCells.filter({ hasText: /^001$/ })).toHaveCount(1);
	await expect(visibleCells.filter({ hasText: /^22$/ })).toHaveCount(1);
	await expect(visibleCells.filter({ hasText: /^333$/ })).toHaveCount(1);
	await goToCell(page, 'A1');
	await typeInActiveCell(page, 'order 99');
	await expect(visibleCells.filter({ hasText: /^99$/ })).toHaveCount(1);
	await expect(visibleCells.filter({ hasText: /^001$/ })).toHaveCount(0);
	await expect(visibleCells.filter({ hasText: /^333$/ })).toHaveCount(0);
	expect(errors).toEqual([]);
});

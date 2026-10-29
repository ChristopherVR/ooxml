import { expect, test, type Page } from '@playwright/test';
import { openFixture } from './fixture-files';
import { clickRibbonControl, editor, FRAMEWORKS, goToCell, pageErrors, ribbon } from './helpers';

// excel-sparklines.xlsx: F1 line, F2 column, F3 win/loss and F4 line, each its own group. Settings
// of a group (type, points, colours) apply to every sparkline in it, so grouping F1:F2 makes the
// two share one type, and ungrouping lets them differ again.
const kinds = (page: Page): Promise<string[]> =>
	editor(page).evaluate((node) => {
		const cells = [...node.shadowRoot!.querySelectorAll<HTMLElement>('.xg-c')]
			.filter((cell) => cell.querySelector(':scope > .xg-spark'))
			.sort((a, b) => Number.parseFloat(a.style.top) - Number.parseFloat(b.style.top));
		return cells.map((cell) => {
			const svg = cell.querySelector(':scope > .xg-spark')!;
			return svg.querySelector('polyline') ? 'line' : 'column';
		});
	});

async function selectRange(page: Page, from: string, to: string) {
	await goToCell(page, from);
	await page.keyboard.down('Shift');
	await goToCell(page, to);
	await page.keyboard.up('Shift');
}

async function command(page: Page, id: string) {
	await ribbon(page).getByRole('tab', { name: 'Insert', exact: true }).click();
	await clickRibbonControl(page, ribbon(page).locator(`[data-command="${id}"]`).first());
}

async function setType(page: Page, cell: string, type: 'line' | 'column') {
	await goToCell(page, cell);
	await command(page, 'sparkline.settings');
	const settings = editor(page).locator('[data-dialog="sparkline-settings"]');
	await settings.getByLabel('Sparkline Type:', { exact: true }).selectOption(type);
	await settings.getByRole('button', { name: 'OK', exact: true }).click();
	await expect(settings).toBeHidden();
}

for (const framework of FRAMEWORKS)
	test(`${framework}: Group and Ungroup Sparklines share and split settings`, async ({ page }) => {
		const errors = pageErrors(page);
		await openFixture(page, framework, 'excel-sparklines.xlsx');
		await expect.poll(() => kinds(page)).toEqual(['line', 'column', 'column', 'line']);

		// Separate groups: changing F1 leaves F2 alone.
		await setType(page, 'F1', 'column');
		await expect.poll(() => kinds(page)).toEqual(['column', 'column', 'column', 'line']);
		await setType(page, 'F1', 'line');
		await expect.poll(() => kinds(page)).toEqual(['line', 'column', 'column', 'line']);

		// Grouped: F2 follows F1.
		await goToCell(page, 'F1');
		await page.keyboard.press('Shift+ArrowDown');
		await command(page, 'sparkline.group');
		await setType(page, 'F1', 'line');
		await expect.poll(() => kinds(page)).toEqual(['line', 'line', 'column', 'line']);

		// Ungrouped: F2 stays as it is when F1 changes.
		await goToCell(page, 'F1');
		await page.keyboard.press('Shift+ArrowDown');
		await command(page, 'sparkline.ungroup');
		await setType(page, 'F1', 'column');
		await expect.poll(() => kinds(page)).toEqual(['column', 'line', 'column', 'line']);
		expect(errors).toEqual([]);
	});

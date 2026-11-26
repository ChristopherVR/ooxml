import { expect, test } from '@playwright/test';
import { clickInEditor, dialogByHeading, insertTableOfSize, newDocument } from './helpers';
import { savedDocument } from './saved-document';

test.use({ viewport: { width: 1700, height: 960 } });

for (const gesture of ['drag', 'Shift+Arrow'] as const) {
	test(`merge two cells selected by ${gesture}, then split them`, async ({ page }) => {
		await page.goto('/?framework=vanilla');
		await newDocument(page);
		const across = gesture === 'Shift+Arrow';
		await insertTableOfSize(page);
		const editor = page.locator('docx-editor');
		const cells = editor.locator('.ProseMirror td');
		const second = cells.nth(across ? 1 : 2);
		await clickInEditor(cells.first().locator('p'));
		await page.keyboard.type('First');
		await clickInEditor(second.locator('p'));
		await page.keyboard.type('Second');
		await clickInEditor(cells.first().locator('p'));
		await editor.getByRole('tab', { name: 'Table Layout', exact: true }).click();
		const merge = editor.getByRole('button', { name: 'Merge cells', exact: true });
		await expect(merge).toBeDisabled();
		await expect(merge).toHaveAccessibleDescription('Select two or more cells that can be merged.');

		if (across) {
			await clickInEditor(cells.first().locator('p'));
			await page.keyboard.press('End');
			await page.keyboard.press('Shift+ArrowRight');
			await page.keyboard.press('Shift+ArrowRight');
		} else {
			const origin = await cells.first().locator('p').boundingBox();
			const destination = await second.locator('p').boundingBox();
			if (!origin || !destination) throw new Error('The cells are outside the editor.');
			await page.mouse.move(origin.x + 1, origin.y + origin.height / 2);
			await page.mouse.down();
			await page.mouse.move(destination.x + 25, destination.y + destination.height / 2, {
				steps: 8,
			});
			await page.mouse.up();
		}
		await expect(merge).toBeEnabled();
		await merge.click();
		await expect(cells.first()).toHaveAttribute(across ? 'colspan' : 'rowspan', '2');
		await expect(cells.first().locator('p')).toHaveText(['First', 'Second']);
		await expect(merge).toBeDisabled();
		const { xml } = await savedDocument(page, 'merged-new-table.docx');
		await expect(cells).toHaveCount(3);
		if (across) expect(xml).toContain('<w:gridSpan w:val="2"');
		else {
			expect(xml).toContain('<w:vMerge w:val="restart"');
			expect(xml).toMatch(/<w:vMerge(?: w:val="continue")?\s*\/>/);
		}

		await editor.getByRole('button', { name: 'Split cells', exact: true }).click();
		const dialog = dialogByHeading(editor, 'Split cells');
		await expect(dialog.getByRole('spinbutton', { name: 'Number of columns' })).toHaveValue(
			across ? '2' : '1',
		);
		await expect(dialog.getByRole('spinbutton', { name: 'Number of rows' })).toHaveValue(
			across ? '1' : '2',
		);
		const ok = dialog.getByRole('button', { name: 'OK', exact: true });
		if (across) {
			const columns = dialog.getByRole('spinbutton', { name: 'Number of columns' });
			await columns.fill('3');
			await expect(ok).toBeDisabled();
			await expect(ok).toHaveAccessibleDescription(
				'These dimensions cannot split the selected cell.',
			);
			await columns.fill('2');
		}
		await ok.click();
		await expect(cells).toHaveCount(4);
		await expect(cells.first()).toHaveText(across ? 'First' : 'FirstSecond');
		if (across) await expect(cells.nth(1)).toHaveText('Second');
		const width = editor.getByRole('spinbutton', { name: 'Column Width (inches)' });
		const height = editor.getByRole('spinbutton', { name: 'Row Height (inches)' });
		await expect(width).toBeEnabled();
		await expect(height).toBeEnabled();
		await width.fill('1.25');
		await width.press('Enter');
		await expect(cells.first()).toHaveAttribute('style', /width:\s*120px/);
		await height.fill('0.75');
		await height.press('Enter');
		await expect(cells.first().locator('..')).toHaveAttribute('style', /height:\s*72px/);
		const split = await savedDocument(page, 'split-new-table.docx');
		expect(split.xml).not.toMatch(/<w:(?:gridSpan|vMerge)\b/);
		expect(split.xml).toMatch(/<w:gridCol w:w="1800"/);
		expect(split.xml).toMatch(/<w:trHeight\b[^>]*w:val="1080"/);
	});
}

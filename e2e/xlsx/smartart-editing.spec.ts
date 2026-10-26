import { expect, test, type Page } from '@playwright/test';
import { FRAMEWORKS, editor, editorProperty, newWorkbook, pageErrors } from './helpers';

type Book = {
	sheets: {
		drawings: {
			kind: string;
			colorsId?: string;
			quickStyleId?: string;
			nodes?: { id: string; text: string; parentId?: string }[];
		}[];
	}[];
};

/** The SmartArt node texts with their depth (two spaces per level), as the text pane shows them. */
async function outline(page: Page): Promise<string[]> {
	const book = await editorProperty<Book>(page, 'workbook');
	const art = book.sheets[0]!.drawings.find((drawing) => drawing.kind === 'smartArt');
	const nodes = art?.nodes ?? [];
	const ids = new Set(nodes.map((node) => node.id));
	const levels = new Map<string, number>();
	return nodes.map((node) => {
		const parent = node.parentId && ids.has(node.parentId) ? node.parentId : undefined;
		const level = parent ? (levels.get(parent) ?? 0) + 1 : 0;
		levels.set(node.id, level);
		return `${'  '.repeat(level)}${node.text}`;
	});
}

const pane = (page: Page) => editor(page).locator('office-ui-smartart-text-pane');
const lines = (page: Page) => pane(page).locator('input[data-node]');

for (const framework of FRAMEWORKS)
	test(`${framework}: inserts SmartArt, edits it through the text pane and Design tab, saves and reloads`, async ({
		page,
	}) => {
		const errors = pageErrors(page);
		await newWorkbook(page, framework);
		await editor(page).getByRole('tab', { name: 'Insert', exact: true }).click();
		await editor(page).getByRole('button', { name: 'SmartArt', exact: true }).first().click();
		const dialog = editor(page).locator('[data-dialog="insert-smartart"]');
		await expect(dialog.locator('office-ui-smartart-picker')).toBeVisible();
		await dialog.getByRole('button', { name: 'OK', exact: true }).click();
		await expect(dialog).toHaveCount(0);

		// The new diagram is drawn (approximately) and selected, with the text pane open on it.
		await expect(editor(page).locator('office-ui-smartart svg')).toBeVisible();
		await expect(
			editor(page).getByRole('tab', { name: 'SmartArt Design', exact: true }),
		).toBeVisible();
		await expect(pane(page)).toBeVisible();
		const empty = await lines(page).count();
		expect(empty).toBeGreaterThan(1);

		// Typing, Enter (add after), Tab (demote) and Shift+Tab (promote) in the text pane.
		await lines(page).first().fill('Alpha');
		await lines(page).first().press('Enter');
		await expect(lines(page)).toHaveCount(empty + 1);
		await lines(page).nth(1).fill('Beta');
		await lines(page).nth(1).press('Tab');
		await expect
			.poll(() => outline(page))
			.toEqual(['Alpha', '  Beta', ...Array(empty - 1).fill('')]);
		await lines(page).nth(1).press('Shift+Tab');
		await expect.poll(() => outline(page)).toEqual(['Alpha', 'Beta', ...Array(empty - 1).fill('')]);

		// Backspace in an empty line removes its node.
		const last = lines(page).last();
		await last.focus();
		await last.press('Backspace');
		await expect(lines(page)).toHaveCount(empty);

		// The SmartArt Design tab: Add Shape After and Demote act on the pane's node (Beta).
		await lines(page).nth(1).focus();
		await editor(page).getByRole('tab', { name: 'SmartArt Design', exact: true }).click();
		await editor(page).getByRole('button', { name: 'Add Shape', exact: true }).first().click();
		await editor(page).getByRole('menuitem', { name: 'Add Shape After', exact: true }).click();
		await expect(lines(page)).toHaveCount(empty + 1);
		await editor(page).getByRole('button', { name: 'Demote', exact: true }).first().click();
		await expect
			.poll(() => outline(page))
			.toEqual(['Alpha', 'Beta', '  ', ...Array(empty - 2).fill('')]);
		await lines(page).nth(2).fill('Gamma');
		const edited = ['Alpha', 'Beta', '  Gamma', ...Array(empty - 2).fill('')];
		await expect.poll(() => outline(page)).toEqual(edited);

		// Undo goes through the workbook's history.
		await editor(page).evaluate((node) => (node as unknown as { undo(): void }).undo());
		await expect.poll(() => outline(page)).not.toEqual(edited);
		await editor(page).evaluate((node) => (node as unknown as { redo(): void }).redo());
		await expect.poll(() => outline(page)).toEqual(edited);
		await editor(page).getByRole('tab', { name: 'SmartArt Design', exact: true }).click();
		await editor(page).getByRole('button', { name: 'SmartArt Styles', exact: true }).click();
		await editor(page).getByRole('button', { name: 'Moderate Effect', exact: true }).click();
		await editor(page).getByRole('button', { name: 'Change Colors', exact: true }).click();
		await editor(page)
			.getByRole('button', { name: 'Colored Fill - Accent 2', exact: true })
			.click();

		const bytes = await editor(page).evaluate(async (node) => [
			...new Uint8Array(
				await (await (node as unknown as { save(): Promise<Blob> }).save()).arrayBuffer(),
			),
		]);
		await editor(page).evaluate(async (node, data) => {
			await (node as unknown as { load(bytes: Uint8Array, name: string): Promise<void> }).load(
				new Uint8Array(data),
				'smartart-roundtrip.xlsx',
			);
		}, bytes);
		await expect.poll(() => outline(page)).toEqual(edited);
		const drawn = editor(page).locator('.xg-smartart').first();
		await expect(drawn).toHaveAttribute('data-rendering', 'interpreted');
		await expect(drawn.locator('office-ui-smartart')).toContainText('Gamma');
		const back = await editorProperty<Book>(page, 'workbook');
		expect(back.sheets[0]!.drawings.find((shape) => shape.kind === 'smartArt')).toMatchObject({
			colorsId: 'urn:ooxml-core/colors/monochromatic2',
			quickStyleId: 'urn:ooxml-core/style/moderate',
		});
		expect(errors).toEqual([]);
	});

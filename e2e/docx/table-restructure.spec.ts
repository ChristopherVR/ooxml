import { expect, test, type Locator, type Page } from '@playwright/test';
import { dialogByHeading, editorSelectionText, fileInput, reopenDocument } from './helpers';
import { savedDocument } from './saved-document';
import { mergedTableFixture } from './support/merged-table-fixture';

test.use({ viewport: { width: 1366, height: 1000 } });

const FRAMEWORKS = ['vanilla', 'react', 'vue', 'angular', 'svelte', 'solid'];

/** A table as the DOM shows it: per row, each cell's text with its spans (1 omitted). */
type Shape = string[][];

const shapeOf = (table: Locator): Promise<Shape> =>
	table.evaluate((node) =>
		Array.from(node.querySelectorAll(':scope > tbody > tr, :scope > tr')).map((row) =>
			Array.from(row.children)
				.filter((cell) => cell.tagName === 'TD' || cell.tagName === 'TH')
				.map((cell) => {
					const text = Array.from(cell.querySelectorAll(':scope > p'))
						.map((p) => (p.textContent ?? '').trim())
						.filter(Boolean)
						.join(' ');
					const col = Number(cell.getAttribute('colspan') ?? 1);
					const rowspan = Number(cell.getAttribute('rowspan') ?? 1);
					return `${text}${col > 1 ? `|c${col}` : ''}${rowspan > 1 ? `|r${rowspan}` : ''}`;
				}),
		),
	);

const surface = (page: Page) => page.locator('docx-editor .ProseMirror').first();
/** The top-level tables: 0 is the merged table, 1 the one that holds the nested table. */
const topTable = (page: Page, index: number) => surface(page).locator('> table').nth(index);
const nestedTable = (page: Page) => surface(page).locator('table table').first();

async function open(page: Page, framework: string): Promise<Locator> {
	page.on('dialog', (dialog) => void dialog.accept());
	await page.goto(`/?framework=${framework}`);
	await (
		await fileInput(page)
	).setInputFiles({
		name: 'merged.docx',
		mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
		buffer: await mergedTableFixture(),
	});
	await expect(surface(page)).toContainText('After');
	return page.locator('docx-editor');
}

/** Selects from the start of the cell showing `from` to the end of `to` (one cell when omitted). */
async function selectCells(editor: Locator, from: string, to = from) {
	await editor.evaluate(
		(element, [a, b]) => {
			/* eslint-disable @typescript-eslint/no-explicit-any */
			const view = (element as any).view;
			let start = -1;
			let end = -1;
			view.state.doc.descendants((node: any, pos: number) => {
				if (node.type.name !== 'paragraph') return;
				if (node.textContent === a && start < 0) start = pos + 1;
				if (node.textContent === b && end < 0) end = pos + 1 + node.content.size;
			});
			if (start < 0 || end < 0) throw new Error(`cells ${a} / ${b} not found`);
			const Selection = view.state.selection.constructor as any;
			view.focus();
			view.dispatch(
				view.state.tr.setSelection(
					Selection.between(view.state.doc.resolve(start), view.state.doc.resolve(end)),
				),
			);
		},
		[from, to],
	);
	await expect.poll(() => editorSelectionText(editor)).toContain(to);
}

async function run(editor: Locator, name: string) {
	await editor.getByRole('tab', { name: 'Table', exact: true }).click();
	const button = editor.getByRole('button', { name, exact: true });
	await expect(button).toBeEnabled();
	await button.click();
}

/** Undo or redo through the keyboard, with the document focused as after any edit. */
async function history(page: Page, editor: Locator, key: 'Control+z' | 'Control+y') {
	await editor.evaluate((element) => (element as any).view.focus());
	await page.keyboard.press(key);
}

/** Saves the document, opens the bytes again and returns the document's reloaded surface. */
async function saveAndReload(page: Page, name: string) {
	const { bytes } = await savedDocument(page, name);
	await reopenDocument(page, name, bytes);
	await expect(surface(page)).toContainText('After');
}

const MERGED: Shape = [
	['A1|c2', 'A3', 'A4'],
	['B1|r2', 'B2', 'B3', 'B4'],
	['C2', 'C3', 'C4'],
];

/**
 * Runs `edit`, then checks the table's shape, that undo restores `before`, redo restores `after`
 * and that the saved document reopens with `after`.
 */
async function checkEdit(
	page: Page,
	editor: Locator,
	table: () => Locator,
	before: Shape,
	after: Shape,
	edit: () => Promise<void>,
	name: string,
) {
	expect(await shapeOf(table())).toEqual(before);
	await edit();
	await expect.poll(() => shapeOf(table())).toEqual(after);
	await history(page, editor, 'Control+z');
	await expect.poll(() => shapeOf(table())).toEqual(before);
	await history(page, editor, 'Control+y');
	await expect.poll(() => shapeOf(table())).toEqual(after);
	await saveAndReload(page, name);
	expect(await shapeOf(table())).toEqual(after);
}

for (const framework of FRAMEWORKS) {
	test.describe(framework, () => {
		test('insert row above a row under a vertical merge', async ({ page }) => {
			const editor = await open(page, framework);
			await selectCells(editor, 'A3');
			await checkEdit(
				page,
				editor,
				() => topTable(page, 0),
				MERGED,
				// The new row has the shape of the row below it, merged cell included.
				[['|c2', '', ''], ...MERGED],
				() => run(editor, 'Insert row above'),
				`${framework}-row-above.docx`,
			);
		});

		test('insert row inside the vertical merge grows it', async ({ page }) => {
			const editor = await open(page, framework);
			await selectCells(editor, 'B2');
			await checkEdit(
				page,
				editor,
				() => topTable(page, 0),
				MERGED,
				[MERGED[0]!, ['B1|r3', 'B2', 'B3', 'B4'], ['', '', ''], MERGED[2]!],
				() => run(editor, 'Insert row below'),
				`${framework}-row-below.docx`,
			);
		});

		test('insert column left of a cell after a horizontal merge', async ({ page }) => {
			const editor = await open(page, framework);
			await selectCells(editor, 'A3');
			await checkEdit(
				page,
				editor,
				() => topTable(page, 0),
				MERGED,
				[
					['A1|c2', '', 'A3', 'A4'],
					['B1|r2', 'B2', '', 'B3', 'B4'],
					['C2', '', 'C3', 'C4'],
				],
				() => run(editor, 'Insert column left'),
				`${framework}-column-left.docx`,
			);
		});

		test('insert column right of a horizontal merge, as Word does', async ({ page }) => {
			const editor = await open(page, framework);
			await selectCells(editor, 'A1');
			// Word measures a single merged cell from its first column: the rows below get the
			// new cell after their first column, the merged row after the merged cell.
			await checkEdit(
				page,
				editor,
				() => topTable(page, 0),
				MERGED,
				[
					['A1|c2', '', 'A3', 'A4'],
					['B1|r2', '', 'B2', 'B3', 'B4'],
					['', 'C2', 'C3', 'C4'],
				],
				() => run(editor, 'Insert column right'),
				`${framework}-column-right.docx`,
			);
		});

		test('delete a row of the vertical merge shrinks it', async ({ page }) => {
			const editor = await open(page, framework);
			await selectCells(editor, 'C2');
			await checkEdit(
				page,
				editor,
				() => topTable(page, 0),
				MERGED,
				[MERGED[0]!, ['B1', 'B2', 'B3', 'B4']],
				() => run(editor, 'Delete row'),
				`${framework}-delete-row.docx`,
			);
		});

		test('delete the last column of a merged table', async ({ page }) => {
			const editor = await open(page, framework);
			await selectCells(editor, 'A4');
			await checkEdit(
				page,
				editor,
				() => topTable(page, 0),
				MERGED,
				[
					['A1|c2', 'A3'],
					['B1|r2', 'B2', 'B3'],
					['C2', 'C3'],
				],
				() => run(editor, 'Delete column'),
				`${framework}-delete-column.docx`,
			);
		});

		test('merge two cells in a row, then split the result again', async ({ page }) => {
			const editor = await open(page, framework);
			await selectCells(editor, 'B3', 'B4');
			await checkEdit(
				page,
				editor,
				() => topTable(page, 0),
				MERGED,
				[MERGED[0]!, ['B1|r2', 'B2', 'B3 B4|c2'], MERGED[2]!],
				// (a merged cell keeps the text of every cell it took in)
				() => run(editor, 'Merge cells'),
				`${framework}-merge.docx`,
			);
			await selectCells(editor, 'B3');
			await run(editor, 'Split cells');
			const dialog = dialogByHeading(editor, 'Split cells');
			await dialog.getByRole('button', { name: 'OK', exact: true }).click();
			await expect
				.poll(() => shapeOf(topTable(page, 0)))
				// Each paragraph of the merged cell goes to a cell of its own.
				.toEqual(MERGED);
		});

		test('merge a block of cells across two rows', async ({ page }) => {
			const editor = await open(page, framework);
			// A text selection from B3 to C3 covers the cells between them in document order too.
			await selectCells(editor, 'B3', 'C3');
			await checkEdit(
				page,
				editor,
				() => topTable(page, 0),
				MERGED,
				[MERGED[0]!, ['B1|r2', 'B2 B3 B4 C2 C3 C4|c3|r2'], []],
				() => run(editor, 'Merge cells'),
				`${framework}-merge-down.docx`,
			);
		});

		test('split a horizontally merged cell into its columns', async ({ page }) => {
			const editor = await open(page, framework);
			await selectCells(editor, 'A1');
			await run(editor, 'Split cells');
			const dialog = dialogByHeading(editor, 'Split cells');
			await dialog.getByRole('spinbutton', { name: 'Number of columns' }).fill('2');
			await dialog.getByRole('button', { name: 'OK', exact: true }).click();
			const split = [['A1', '', 'A3', 'A4'], ...MERGED.slice(1)];
			await expect.poll(() => shapeOf(topTable(page, 0))).toEqual(split);
			await saveAndReload(page, `${framework}-split.docx`);
			expect(await shapeOf(topTable(page, 0))).toEqual(split);
		});

		test('a nested table is edited in place', async ({ page }) => {
			const editor = await open(page, framework);
			const NESTED: Shape = [
				['N11', 'N12'],
				['N21', 'N22'],
			];
			expect(await shapeOf(nestedTable(page))).toEqual(NESTED);
			await selectCells(editor, 'N12');
			await checkEdit(
				page,
				editor,
				() => nestedTable(page),
				NESTED,
				[NESTED[0]!, ['', ''], NESTED[1]!],
				() => run(editor, 'Insert row below'),
				`${framework}-nested-row.docx`,
			);
			await selectCells(editor, 'N11', 'N12');
			await run(editor, 'Merge cells');
			await expect
				.poll(() => shapeOf(nestedTable(page)))
				.toEqual([['N11 N12|c2'], ['', ''], NESTED[1]!]);
			await selectCells(editor, 'N21');
			await run(editor, 'Delete row');
			const finalShape: Shape = [['N11 N12|c2'], ['', '']];
			await expect.poll(() => shapeOf(nestedTable(page))).toEqual(finalShape);
			await saveAndReload(page, `${framework}-nested.docx`);
			expect(await shapeOf(nestedTable(page))).toEqual(finalShape);
			// The tables around it are untouched.
			expect(await shapeOf(topTable(page, 0))).toEqual(MERGED);
			await expect(surface(page)).toContainText('Tail');
		});
	});
}

for (const framework of FRAMEWORKS) {
	test.describe(`${framework} ribbon`, () => {
		test('table ribbon commands are enabled on a merged table', async ({ page }) => {
			const editor = await open(page, framework);
			expect(await shapeOf(topTable(page, 0))).toEqual(MERGED);
			const enabled = async (names: string[]) => {
				await editor.getByRole('tab', { name: 'Table', exact: true }).click();
				for (const name of names)
					await expect(editor.getByRole('button', { name, exact: true })).toBeEnabled();
			};
			await selectCells(editor, 'A4');
			await enabled([
				'Insert row above',
				'Insert row below',
				'Insert column left',
				'Insert column right',
				'Delete row',
				'Delete column',
				'Split cells',
				'Delete table',
			]);
			await selectCells(editor, 'B3', 'B4');
			await enabled(['Merge cells']);
			// A cell that is part of a merge can still be split and have rows added around it.
			await selectCells(editor, 'A1');
			await enabled(['Split cells', 'Insert row below']);
		});
	});
}

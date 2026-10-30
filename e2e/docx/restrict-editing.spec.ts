import { expect, test, type Page } from '@playwright/test';
import type { EditorView } from 'prosemirror-view';
import { newDocument } from './helpers';

// Review > Restrict Editing: a read only document refuses typing outside its editable region and
// accepts it inside, and the docked pane stops the protection with its password.
const texts = (page: Page) =>
	page.locator('docx-editor').evaluate((element) => {
		const { doc } = (element as unknown as { view: EditorView }).view.state;
		const result: string[] = [];
		doc.forEach((node) => result.push(node.textContent));
		return result;
	});

async function protect(page: Page) {
	await page.locator('docx-editor').evaluate((element) => {
		const editor = element as unknown as { documentModel: Record<string, unknown> };
		const model = editor.documentModel;
		model.blocks = [
			{
				type: 'paragraph',
				id: 'p1',
				runs: [{ text: 'hello world' }],
				permissions: [{ id: '1', group: 'everyone', from: 6, to: 11 }],
			},
			{ type: 'paragraph', id: 'p2', runs: [{ text: 'locked' }] },
		];
		model.protection = { edit: 'readOnly', enforced: true };
		editor.documentModel = model;
	});
	await expect.poll(() => texts(page)).toEqual(['hello world', 'locked']);
}

/** Puts the caret at `offset` of the paragraph at `index`, through the editor's own selection. */
const placeCaret = (page: Page, index: number, offset: number) =>
	page.locator('docx-editor').evaluate(
		async (element, [at, inside]) => {
			const view = (element as unknown as { view: EditorView }).view;
			let pos = 0;
			for (let i = 0; i < at!; i++) pos += view.state.doc.child(i).nodeSize;

			view.dispatch(
				view.state.tr.setSelection(
					(
						view.state.selection.constructor as unknown as {
							create: (d: unknown, a: number) => never;
						}
					).create(view.state.doc, pos + 1 + inside!),
				),
			);
			view.focus();
		},
		[index, offset],
	);

test('a read only document accepts typing only inside its editable region', async ({ page }) => {
	const errors: string[] = [];
	page.on('pageerror', (error) => errors.push(error.message));
	await page.goto('/?framework=vanilla');
	await newDocument(page);
	await protect(page);

	await placeCaret(page, 0, 2);
	await page.keyboard.type('XYZ');
	await placeCaret(page, 1, 2);
	await page.keyboard.type('XYZ');
	expect(await texts(page)).toEqual(['hello world', 'locked']);

	await placeCaret(page, 0, 11);
	await page.keyboard.type('!!');
	expect(await texts(page)).toEqual(['hello world!!', 'locked']);
	await expect(page.locator('docx-editor .dve-editable-region')).toContainText('world!!');
	expect(errors).toEqual([]);
});

test('the Restrict Editing pane docks on the right and stops the protection', async ({ page }) => {
	await page.setViewportSize({ width: 1800, height: 900 });
	await page.goto('/?framework=vanilla');
	await newDocument(page);
	await protect(page);
	const editor = page.locator('docx-editor');
	await editor.getByRole('tab', { name: 'Review', exact: true }).click();
	await editor.getByRole('button', { name: 'Restrict editing' }).click();
	const pane = editor.locator('.dve-restrict-pane');
	await expect(pane).toBeVisible();
	const box = (await pane.boundingBox())!;
	const canvas = (await editor.locator('.ProseMirror').boundingBox())!;
	expect(box.x).toBeGreaterThan(canvas.x);

	await pane.getByRole('button', { name: 'Find next region I can edit' }).click();
	await expect
		.poll(() =>
			editor.evaluate((element) => {
				const { state } = (element as unknown as { view: EditorView }).view;
				return state.doc.textBetween(state.selection.from, state.selection.to);
			}),
		)
		.toBe('world');
	await pane.getByRole('button', { name: 'Stop protection' }).click();
	await expect(pane.getByRole('button', { name: 'Yes, start enforcing protection' })).toBeVisible();
	await placeCaret(page, 1, 2);
	await page.keyboard.type('ok');
	expect(await texts(page)).toEqual(['hello world', 'lookcked']);
});

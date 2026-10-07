import { test, expect } from '@playwright/test';
import type { DocxEditorElement } from '../../viewers/docx/packages/web-component/src';
import { newDocument } from './helpers';

for (const framework of ['vanilla', 'react', 'vue', 'angular', 'svelte', 'solid'])
	test(`${framework}: review modes filter inline objects and retain their revision history`, async ({
		page,
	}) => {
		await page.setViewportSize({ width: 2400, height: 1000 });
		await page.goto(`/?framework=${framework}`);
		await newDocument(page);
		const editor = page.locator('docx-editor');
		await editor.evaluate((element) => {
			const host = element as DocxEditorElement;
			host.documentModel = {
				...host.documentModel!,
				blocks: [
					{
						type: 'paragraph',
						id: 'p',
						runs: [
							{ text: 'Keep' },
							{
								text: '',
								image: {
									relId: 'rId1',
									partName: 'word/media/image.png',
									contentType: 'image/png',
									widthPx: 20,
									heightPx: 20,
								},
								revision: { id: 'picture', kind: 'insert', author: 'Ada' },
							},
							{ text: '', break: 'page', revision: { id: 'break', kind: 'delete', author: 'Ada' } },
							{
								text: '',
								noteReference: { kind: 'footnote', id: '1' },
								revision: { id: 'note', kind: 'moveFrom', author: 'Ada', move: { name: 'move' } },
							},
							{
								text: '',
								equation: {
									omml: '<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:r><m:t>x</m:t></m:r></m:oMath>',
									display: false,
								},
								revision: { id: 'equation', kind: 'moveTo', author: 'Ada', move: { name: 'move' } },
							},
						],
					},
				],
			};
		});
		const body = editor.locator('.dve-paper > .ProseMirror');
		await expect(body).toContainText('Keep');
		const source = await editor.evaluate((element) => (element as DocxEditorElement).documentModel);
		await editor.getByRole('tab', { name: 'Review', exact: true }).click();
		const mode = editor.getByRole('combobox', { name: 'Display for review', exact: true });
		const inserted = ['.dve-picture', '.dve-equation'];
		const deleted = ['.dve-break-marker', '.dve-note-reference'];
		for (const value of ['original', 'final', 'simple', 'all']) {
			await mode.selectOption(value, { force: true });
			for (const selector of [...inserted, ...deleted])
				if (
					inserted.includes(selector)
						? value === 'original'
						: value === 'final' || value === 'simple'
				)
					await expect(body.locator(selector)).toBeHidden();
				else await expect(body.locator(selector)).toBeVisible();
			expect(
				await editor.evaluate((element) => (element as DocxEditorElement).documentModel),
			).toEqual(source);
		}
		await body.click({ position: { x: 10, y: 10 } });
		await page.keyboard.press('Control+End');
		await page.keyboard.type('!');
		const next = await editor.evaluate((element) => (element as DocxEditorElement).documentModel);
		const paragraph = next!.blocks[0]!;
		if (paragraph.type !== 'paragraph') throw new Error('Expected paragraph');
		expect(
			paragraph.runs
				.filter((run) => run.image || run.break || run.noteReference || run.equation)
				.map((run) => run.revision!.kind),
		).toEqual(['insert', 'delete', 'moveFrom', 'moveTo']);
	});

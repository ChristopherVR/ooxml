import { readFile } from 'node:fs/promises';
import { test, expect } from '@playwright/test';
import { loadDocx, saveDocx } from '../../viewers/docx/packages/core/src/index';
import type { DocxEditorElement } from '../../viewers/docx/packages/web-component/src';
import { fileInput } from './helpers';

async function fixture(suffix: string) {
	const run = (
		await loadDocx(
			await readFile(
				new URL(
					`../../src/core/docx/__fixtures__/review-formatting/multiple-${suffix}.docx`,
					import.meta.url,
				),
			),
		)
	).model;
	const paragraph = (
		await loadDocx(
			await readFile(
				new URL(
					`../../src/core/docx/__fixtures__/review-paragraph-formatting/multiple-${suffix}.docx`,
					import.meta.url,
				),
			),
		)
	).model;
	const blocks = [run.blocks[0]!, paragraph.blocks[0]!];
	run.blocks = [{ type: 'paragraph', id: 'body', runs: [{ text: 'Story review fixture' }] }];
	run.sections![0]!.endsAtBlockId = 'body';
	run.sections![0]!.headers = { default: { partName: 'word/header1.xml', blocks } };
	run.sections![0]!.footers = { default: { partName: 'word/footer1.xml', blocks } };
	run.footnotes = [{ id: '1', blocks }];
	return Buffer.from(await saveDocx(run));
}

for (const framework of ['vanilla', 'react', 'vue', 'angular', 'svelte', 'solid'])
	test(`${framework}: Original projects native formatting in story previews and active editors`, async ({
		page,
	}) => {
		await page.setViewportSize({ width: 2400, height: 1000 });
		await page.goto(`/?framework=${framework}`);
		const input = await fileInput(page);
		const editor = page.locator('docx-editor');
		const previews = editor.locator('.dve-header-footer-body, .dve-note-body');
		const appearance = () =>
			previews.evaluateAll((elements) =>
				elements.map((element) =>
					[...element.querySelectorAll('p')].map((p) => ({
						text: p.textContent,
						align: getComputedStyle(p).textAlign,
						before: getComputedStyle(p).marginTop,
						after: getComputedStyle(p).marginBottom,
						left: getComputedStyle(p).marginLeft,
						indent: getComputedStyle(p).textIndent,
						runs: [...p.querySelectorAll('span')]
							.filter((span) => span.firstChild?.nodeType === Node.TEXT_NODE)
							.map((span) => {
								const style = getComputedStyle(span);
								return {
									text: span.textContent,
									font: style.fontFamily,
									size: style.fontSize,
									weight: style.fontWeight,
									italic: style.fontStyle,
									color: style.color,
								};
							}),
					})),
				),
			);
		await input.setInputFiles({
			name: 'before.docx',
			mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
			buffer: await fixture('before'),
		});
		await expect(editor.locator('.ProseMirror')).toContainText('Story review fixture');
		await expect(previews).toHaveCount(3);
		const before = await appearance();
		await input.setInputFiles({
			name: 'tracked.docx',
			mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
			buffer: await fixture('tracked'),
		});
		await expect.poll(appearance).not.toEqual(before);
		const source = await editor.evaluate((element) => (element as DocxEditorElement).documentModel);
		await editor.getByRole('tab', { name: 'Review', exact: true }).click();
		const mode = editor.getByRole('combobox', { name: 'Display for review', exact: true });
		await mode.selectOption('original', { force: true });
		await expect.poll(appearance).toEqual(before);
		const header = editor.locator('.dve-header [data-slot="default"]');
		await header.dblclick();
		const active = header.locator('.ProseMirror');
		await expect(active).toHaveAttribute('data-review-display', 'original');
		await editor.getByRole('tab', { name: 'Review', exact: true }).click();
		await mode.selectOption('final', { force: true });
		await expect(active).toHaveAttribute('data-review-display', 'final');
		await mode.selectOption('original', { force: true });
		await expect(active).toHaveAttribute('data-review-display', 'original');
		await active.press('Escape');
		await expect(header.locator('.ProseMirror')).toHaveCount(0);
		await expect.poll(appearance).toEqual(before);
		expect(
			await editor.evaluate((element) => (element as DocxEditorElement).documentModel),
		).toEqual(source);
	});

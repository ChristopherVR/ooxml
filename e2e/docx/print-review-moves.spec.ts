import { fileURLToPath } from 'node:url';
import { test, expect } from '@playwright/test';
import JSZip from 'jszip';
import type { EditorView } from 'prosemirror-view';
import { fileInput } from './helpers';
import type { DocxEditorElement } from '../../viewers/docx/packages/web-component/src';

const fixture = fileURLToPath(
	new URL('../../src/core/docx/__fixtures__/review-moves/word-saved.docx', import.meta.url),
);
for (const framework of ['vanilla', 'react', 'vue', 'angular', 'svelte', 'solid'])
	test(`${framework}: Print Layout displays native move sources and destinations without resolving revisions`, async ({
		page,
	}) => {
		await page.setViewportSize({ width: 2400, height: 1000 });
		const errors: string[] = [];
		page.on('pageerror', (error) => errors.push(error.message));
		await page.goto(`/?framework=${framework}`);
		await (await fileInput(page)).setInputFiles(fixture);
		const editor = page.locator('docx-editor');
		await expect(editor.locator('.ProseMirror')).toContainText('SecondMove');
		const source = await editor.evaluate(
			(element) => (element as DocxEditorElement).documentModel!,
		);
		await editor.getByRole('tab', { name: 'View', exact: true }).click();
		await editor.getByRole('combobox', { name: 'Layout view', exact: true }).selectOption('print');
		const blocks = editor.locator(
			'.dve-canvas > .dve-print-pages .dve-print-column .dve-print-block',
		);
		await expect(blocks).toHaveCount(3);
		await expect
			.poll(async () => (await blocks.allTextContents()).map((text) => text.trim()))
			.toEqual(['FirstMove', 'SecondMove', 'Target FirstMove SecondMove']);
		await editor.getByRole('tab', { name: 'Review', exact: true }).click();
		const mode = editor.getByRole('combobox', { name: 'Display for review', exact: true });
		await mode.selectOption('original', { force: true });
		await expect
			.poll(async () => (await blocks.allTextContents()).map((text) => text.trim()))
			.toEqual(['FirstMove', 'SecondMove', 'Target']);
		await mode.selectOption('final', { force: true });
		await expect
			.poll(async () => (await blocks.allTextContents()).map((text) => text.trim()))
			.toEqual(['', '', 'Target FirstMove SecondMove']);
		await mode.selectOption('simple', { force: true });
		await expect
			.poll(async () => (await blocks.allTextContents()).map((text) => text.trim()))
			.toEqual(['', '', 'Target FirstMove SecondMove']);
		await mode.selectOption('original', { force: true });
		await expect
			.poll(async () => (await blocks.allTextContents()).map((text) => text.trim()))
			.toEqual(['FirstMove', 'SecondMove', 'Target']);
		await blocks.nth(1).locator('span').filter({ hasText: 'SecondMove' }).click();
		expect(
			await editor.evaluate(
				(element) =>
					(element as unknown as { view: EditorView }).view.state.selection.$from.parent.attrs.id,
			),
		).toBe(source.blocks[1]!.id);
		expect(
			await editor.evaluate((element) => (element as DocxEditorElement).documentModel),
		).toEqual(source);
		const bytes = await editor.evaluate(async (element) =>
			Array.from(await (element as DocxEditorElement).saveBytes()),
		);
		const xml = await (
			await JSZip.loadAsync(new Uint8Array(bytes))
		)
			.file('word/document.xml')!
			.async('string');
		expect(xml).toContain('moveFrom');
		expect(xml).toContain('moveTo');
		expect(errors).toEqual([]);
	});

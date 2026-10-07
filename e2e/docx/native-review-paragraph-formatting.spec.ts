import { fileURLToPath } from 'node:url';
import { expect, test } from '@playwright/test';
import JSZip from 'jszip';
import { fileInput } from './helpers';
import type { DocxEditorElement } from '../../viewers/docx/packages/web-component/src';

const fixture = fileURLToPath(
	new URL(
		'../../src/core/docx/__fixtures__/review-paragraph-formatting/multiple-tracked.docx',
		import.meta.url,
	),
);
for (const framework of ['vanilla', 'react', 'vue', 'angular', 'svelte', 'solid']) {
	for (const mode of ['Accept', 'Reject'] as const)
		test(`${framework}: ${mode.toLowerCase()} native paragraph formatting with undo and export`, async ({
			page,
		}) => {
			await page.setViewportSize({ width: 2400, height: 1000 });
			const errors: string[] = [];
			page.on('pageerror', (error) => errors.push(error.message));
			await page.goto(`/?framework=${framework}`);
			await (await fileInput(page)).setInputFiles(fixture);
			const editor = page.locator('docx-editor');
			const body = editor.locator('.ProseMirror');
			await expect(body).toContainText('Paragraph formatting');
			await editor.getByRole('tab', { name: 'Review', exact: true }).click();
			await editor.getByRole('button', { name: 'Next change', exact: true }).click();
			await editor.getByRole('button', { name: mode, exact: true }).click();
			const paragraph = await editor.evaluate(
				(element) => (element as DocxEditorElement).documentModel!.blocks[0]!,
			);
			if (paragraph.type !== 'paragraph') throw new Error('Expected paragraph');
			expect(paragraph.formatRevision).toBeUndefined();
			expect(paragraph.align ?? 'left').toBe(mode === 'Reject' ? 'left' : 'center');
			expect(paragraph.spacingBeforeTwips ?? 0).toBe(mode === 'Reject' ? 0 : 360);
			const bytes = await editor.evaluate(async (element) =>
				Array.from(await (element as DocxEditorElement).saveBytes()),
			);
			const xml = await (
				await JSZip.loadAsync(new Uint8Array(bytes))
			)
				.file('word/document.xml')!
				.async('string');
			expect(xml).not.toContain('pPrChange');
			await body.click();
			await page.keyboard.press('Control+z');
			await expect
				.poll(() =>
					editor.evaluate((element) => {
						const paragraph = (element as DocxEditorElement).documentModel!.blocks[0]!;
						return paragraph.type === 'paragraph' && paragraph.formatRevision?.kind;
					}),
				)
				.toBe('paragraphChange');
			await page.keyboard.press('Control+Shift+z');
			await expect
				.poll(() =>
					editor.evaluate((element) => {
						const paragraph = (element as DocxEditorElement).documentModel!.blocks[0]!;
						return paragraph.type === 'paragraph' && Boolean(paragraph.formatRevision);
					}),
				)
				.toBe(false);
			expect(errors).toEqual([]);
		});
}
for (const framework of ['vanilla', 'react', 'vue', 'angular', 'svelte', 'solid'])
	for (const tracked of [true, false])
		test(`${framework}: text editing retains native paragraph history with tracking ${tracked}`, async ({
			page,
		}) => {
			await page.setViewportSize({ width: 2400, height: 1000 });
			const errors: string[] = [];
			page.on('pageerror', (error) => errors.push(error.message));
			await page.goto(`/?framework=${framework}`);
			await (await fileInput(page)).setInputFiles(fixture);
			const editor = page.locator('docx-editor');
			const body = editor.locator('.ProseMirror');
			await expect(body).toContainText('Paragraph formatting');
			const revision = await editor.evaluate((element) => {
				const paragraph = (element as DocxEditorElement).documentModel!.blocks[0]!;
				if (paragraph.type !== 'paragraph') throw new Error('Expected paragraph');
				return paragraph.formatRevision;
			});
			expect(revision!.previousParagraphPropertiesXml).toContain('pPr');
			await editor.getByRole('tab', { name: 'Review', exact: true }).click();
			if (!tracked)
				await editor.getByRole('button', { name: 'Track changes', exact: true }).click();
			await body.click();
			await page.keyboard.press('Control+End');
			await page.keyboard.type('!');
			await expect(body).toContainText('Paragraph formatting!');
			const actual = await editor.evaluate((element) => {
				const paragraph = (element as DocxEditorElement).documentModel!.blocks[0]!;
				if (paragraph.type !== 'paragraph') throw new Error('Expected paragraph');
				return paragraph.formatRevision;
			});
			expect(actual).toEqual(revision);
			const bytes = await editor.evaluate(async (element) =>
				Array.from(await (element as DocxEditorElement).saveBytes()),
			);
			const xml = await (
				await JSZip.loadAsync(new Uint8Array(bytes))
			)
				.file('word/document.xml')!
				.async('string');
			expect(xml).toContain('pPrChange');
			expect(xml).toContain('Paragraph formatting');
			expect(xml.includes('<w:ins')).toBe(tracked);
			expect(xml).toContain('w:cs="Arial"');
			expect(errors).toEqual([]);
		});

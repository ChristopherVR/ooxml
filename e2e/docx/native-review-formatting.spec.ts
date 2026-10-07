import { fileURLToPath } from 'node:url';
import { expect, test } from '@playwright/test';
import JSZip from 'jszip';
import { fileInput } from './helpers';
import type { DocxEditorElement } from '../../viewers/docx/packages/web-component/src';

const fixture = fileURLToPath(
	new URL(
		'../../src/core/docx/__fixtures__/review-formatting/multiple-tracked.docx',
		import.meta.url,
	),
);
for (const framework of ['vanilla', 'react', 'vue', 'angular', 'svelte', 'solid']) {
	for (const mode of ['Accept', 'Reject'] as const)
		test(`${framework}: ${mode.toLowerCase()} native formatting with undo and export`, async ({
			page,
		}) => {
			await page.setViewportSize({ width: 2400, height: 1000 });
			const errors: string[] = [];
			page.on('pageerror', (error) => errors.push(error.message));
			await page.goto(`/?framework=${framework}`);
			await (await fileInput(page)).setInputFiles(fixture);
			const editor = page.locator('docx-editor');
			const marker = editor.locator('.ProseMirror .dve-revision-format');
			await expect(marker).toHaveCount(1);
			await editor.getByRole('tab', { name: 'Review', exact: true }).click();
			await marker.dblclick();
			await editor.getByRole('button', { name: mode, exact: true }).click();
			await expect(marker).toHaveCount(0);
			const runs = await editor.evaluate((element) => {
				const paragraph = (element as DocxEditorElement).documentModel!.blocks[0]!;
				if (paragraph.type !== 'paragraph') throw new Error('Expected paragraph');
				return paragraph.runs;
			});
			expect(runs[0]).toMatchObject({
				text: 'Format me',
				color: mode === 'Reject' ? '#0000FF' : '#008000',
			});
			expect(Boolean(runs[0]!.bold)).toBe(mode === 'Reject');
			expect(runs[0]!.revision).toBeUndefined();
			const bytes = await editor.evaluate(async (element) =>
				Array.from(await (element as DocxEditorElement).saveBytes()),
			);
			const xml = await (
				await JSZip.loadAsync(new Uint8Array(bytes))
			)
				.file('word/document.xml')!
				.async('string');
			expect(xml).not.toContain('rPrChange');
			expect(xml).toContain('Format me');
			await editor.locator('.ProseMirror').click();
			await page.keyboard.press('Control+z');
			await expect(marker).toHaveCount(1);
			await page.keyboard.press('Control+Shift+z');
			await expect(marker).toHaveCount(0);
			expect(errors).toEqual([]);
		});
	for (const tracked of [true, false])
		test(`${framework}: text editing preserves native formatting history with tracking ${tracked}`, async ({
			page,
		}) => {
			await page.setViewportSize({ width: 2400, height: 1000 });
			const errors: string[] = [];
			page.on('pageerror', (error) => errors.push(error.message));
			await page.goto(`/?framework=${framework}`);
			await (await fileInput(page)).setInputFiles(fixture);
			const editor = page.locator('docx-editor');
			const body = editor.locator('.ProseMirror');
			await expect(body).toContainText('Format me');
			await editor.getByRole('tab', { name: 'Review', exact: true }).click();
			if (!tracked)
				await editor.getByRole('button', { name: 'Track changes', exact: true }).click();
			await expect
				.poll(() =>
					editor.evaluate((element) => (element as DocxEditorElement).documentModel!.trackChanges),
				)
				.toBe(tracked);
			await body.click();
			await page.keyboard.press('Control+End');
			await page.keyboard.type('!');
			await expect(body).toContainText('Format me!');
			const bytes = await editor.evaluate(async (element) =>
				Array.from(await (element as DocxEditorElement).saveBytes()),
			);
			const xml = await (
				await JSZip.loadAsync(new Uint8Array(bytes))
			)
				.file('word/document.xml')!
				.async('string');
			expect(xml).toContain('rPrChange');
			expect(xml).toContain('Format me');
			expect(xml.includes('<w:ins')).toBe(tracked);
			expect(xml).toMatch(/<w:rPrChange[^>]*>[\s\S]*?<w:color w:val="0000FF"/);
			expect(errors).toEqual([]);
		});
}

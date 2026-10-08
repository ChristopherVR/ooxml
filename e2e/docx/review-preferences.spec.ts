import { fileURLToPath } from 'node:url';
import { test, expect } from '@playwright/test';
import JSZip from 'jszip';
import type { EditorView } from 'prosemirror-view';
import type { DocxEditorElement } from '../../viewers/docx/packages/web-component/src';
import { fileInput } from './helpers';

const fixture = fileURLToPath(
	new URL('../../src/core/docx/__fixtures__/review-preferences/preferences.docx', import.meta.url),
);
for (const framework of ['vanilla', 'react', 'vue', 'angular', 'svelte', 'solid'])
	test(`${framework}: honors native disabled formatting tracking and exports preferences`, async ({
		page,
	}) => {
		await page.setViewportSize({ width: 2400, height: 1000 });
		const errors: string[] = [];
		page.on('pageerror', (error) => errors.push(error.message));
		await page.goto(`/?framework=${framework}`);
		await (await fileInput(page)).setInputFiles(fixture);
		const editor = page.locator('docx-editor');
		const body = editor.locator('.ProseMirror');
		await expect(body).toContainText('Preference text');
		await body.click();
		await page.keyboard.press('Control+Home');
		const selection = () =>
			editor.evaluate((element) => {
				const view = (element as unknown as { view: EditorView }).view;
				return { from: view.state.selection.from, to: view.state.selection.to };
			});
		// Each press extends the native selection, which ProseMirror reads back on `selectionchange`.
		// Pressing on before it has read the last one can redraw the older selection over the newer
		// native one and lose a step, so wait for every step.
		await expect.poll(selection).toEqual({ from: 1, to: 1 });
		for (let index = 1; index <= 10; index++) {
			await page.keyboard.press('Shift+ArrowRight');
			await expect.poll(selection).toEqual({ from: 1, to: 1 + index });
		}
		await page.keyboard.press('Control+b');
		await expect(body.locator('.dve-revision-format')).toHaveCount(0);
		const firstRun = await editor.evaluate((element) => {
			const paragraph = (element as DocxEditorElement).documentModel!.blocks[0]!;
			if (paragraph.type !== 'paragraph') throw new Error('Expected paragraph');
			return paragraph.runs[0];
		});
		expect(firstRun?.bold).toBe(true);
		await page.keyboard.press('Control+End');
		await page.keyboard.type('!');
		await expect(body).toContainText('Preference text!');
		const bytes = await editor.evaluate(async (element) =>
			Array.from(await (element as DocxEditorElement).saveBytes()),
		);
		const zip = await JSZip.loadAsync(new Uint8Array(bytes));
		const xml = await zip.file('word/document.xml')!.async('string');
		expect(xml).toContain('<w:ins ');
		expect(xml).not.toContain('rPrChange');
		const settings = await zip.file('word/settings.xml')!.async('string');
		expect(settings).toContain('doNotTrackFormatting');
		expect(settings).toContain('doNotTrackMoves');
		await page.keyboard.press('Control+z');
		await expect(body).toHaveText('Preference text');
		expect(errors).toEqual([]);
	});

import { fileURLToPath } from 'node:url';
import { test, expect } from '@playwright/test';
import JSZip from 'jszip';
import { fileInput } from './helpers';
import type { DocxEditorElement } from '../../viewers/docx/packages/web-component/src';
import type { EditorView } from 'prosemirror-view';

const fixture = fileURLToPath(
	new URL('../../src/core/docx/__fixtures__/formatting-actions/baseline.docx', import.meta.url),
);
for (const framework of ['vanilla', 'react', 'vue', 'angular', 'svelte', 'solid'])
	test(`${framework}: records and rejects formatting with export and undo`, async ({ page }) => {
		await page.setViewportSize({ width: 2400, height: 1000 });
		const errors: string[] = [];
		page.on('pageerror', (error) => errors.push(error.message));
		await page.goto(`/?framework=${framework}`);
		await (await fileInput(page)).setInputFiles(fixture);
		const editor = page.locator('docx-editor');
		const body = editor.locator('.ProseMirror');
		await expect(body).toContainText('Format me');
		await editor.getByRole('tab', { name: 'Review', exact: true }).click();
		const track = editor.getByRole('button', { name: 'Track changes', exact: true });
		await track.click();
		await expect(track).toHaveAttribute('aria-pressed', 'true');
		const selection = () =>
			editor.evaluate((element) => {
				const view = (element as unknown as { view: EditorView }).view;
				return { from: view.state.selection.from, to: view.state.selection.to };
			});
		// On a slow runner a binding can still be re-rendering after the toggle and drop some of
		// the key presses, so select again until the first six characters are selected.
		await expect(async () => {
			await body.click();
			await page.keyboard.press('Control+Home');
			for (let index = 0; index < 6; index++) await page.keyboard.press('Shift+ArrowRight');
			expect(await selection()).toEqual({ from: 1, to: 7 });
		}).toPass({ timeout: 15_000 });
		await page.keyboard.press('Control+b');
		const marker = body.locator('.dve-revision-format');
		await expect(marker).toHaveCount(1);
		const bytes = await editor.evaluate(async (element) =>
			Array.from(await (element as DocxEditorElement).saveBytes()),
		);
		const xml = await (
			await JSZip.loadAsync(new Uint8Array(bytes))
		)
			.file('word/document.xml')!
			.async('string');
		expect(xml).toContain('rPrChange');
		expect(xml).toContain('dateUtc');
		await page.keyboard.press('Control+z');
		await expect(marker).toHaveCount(0);
		await page.keyboard.press('Control+Shift+z');
		await expect(marker).toHaveCount(1);
		await marker.dblclick();
		await editor.getByRole('button', { name: 'Reject', exact: true }).click();
		await expect(marker).toHaveCount(0);
		const runs = await editor.evaluate((element) => {
			const paragraph = (element as DocxEditorElement).documentModel!.blocks[0]!;
			if (paragraph.type !== 'paragraph') throw new Error('Expected paragraph');
			return paragraph.runs;
		});
		expect(runs.map((run) => run.text).join('')).toBe('Format me');
		for (const run of runs) expect(Boolean(run.bold)).toBe(false);
		expect(errors).toEqual([]);
	});

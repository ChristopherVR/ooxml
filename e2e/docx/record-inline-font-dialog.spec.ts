import { test, expect } from '@playwright/test';
import { fileURLToPath } from 'node:url';
import { readFile } from 'node:fs/promises';
import type { EditorView } from 'prosemirror-view';
import type { DocxEditorElement } from '../../viewers/docx/packages/web-component/src';
import { fileInput, dialogByHeading, reveal } from './helpers';
import {
	loadDocx,
	type DocumentModel,
	type TextRun,
} from '../../viewers/docx/packages/core/src/index';

function content(model: DocumentModel) {
	return model.blocks.map((block) => {
		if (block.type !== 'paragraph') return '[table]';
		const merged: TextRun[] = [];
		for (const {
			sourceRunPropertiesXml: _source,
			restoredRunPropertiesXml: _restored,
			...run
		} of block.runs) {
			const previous = merged.at(-1);
			const { text, ...properties } = run;
			const { text: _previousText, ...priorProperties } = previous ?? { text: '' };
			if (previous && text && JSON.stringify(properties) === JSON.stringify(priorProperties))
				previous.text += text;
			else merged.push({ ...run });
		}
		return merged;
	});
}

for (const framework of ['vanilla', 'react', 'vue', 'angular', 'svelte', 'solid'])
	for (const name of ['picture', 'note', 'break', 'line-break'])
		test(`${framework}: one Font-dialog edit records native ${name} properties and preserves undo`, async ({
			page,
		}) => {
			await page.setViewportSize({ width: 2400, height: 1000 });
			await page.goto(`/?framework=${framework}`);
			const fixture = fileURLToPath(
				new URL(
					`../../src/core/docx/__fixtures__/review-advanced-object-formatting/${name}-before.docx`,
					import.meta.url,
				),
			);
			await (await fileInput(page)).setInputFiles(fixture);
			const editor = page.locator('docx-editor');
			const body = editor.locator('.dve-paper > .ProseMirror');
			await expect(body).toContainText('Before');
			await editor.getByRole('tab', { name: 'Review', exact: true }).click();
			await editor.getByRole('button', { name: 'Track changes', exact: true }).click();
			const tree = () =>
				editor.evaluate((el) => (el as unknown as { view: EditorView }).view.state.doc.toJSON());
			const initial = await tree();
			await body.click();
			await body.press('Control+a');
			await editor.getByRole('tab', { name: 'Home', exact: true }).click();
			const launcher = editor.getByRole('button', { name: 'Font settings', exact: true });
			await reveal(editor, launcher);
			await launcher.click();
			const dialog = dialogByHeading(editor, 'Font');
			await dialog.getByRole('textbox', { name: 'Size', exact: true }).fill('18');
			await dialog.getByLabel('Font color', { exact: true }).fill('#c00000');
			await dialog.getByRole('checkbox', { name: 'Small caps', exact: true }).check();
			await dialog.getByRole('tab', { name: 'Advanced', exact: true }).click();
			await dialog.getByRole('spinbutton', { name: 'Scale (%)', exact: true }).fill('150');
			await dialog
				.getByRole('combobox', { name: 'Character spacing', exact: true })
				.selectOption('expanded');
			await dialog.getByRole('spinbutton', { name: 'By', exact: true }).fill('1.5');
			await dialog.getByRole('combobox', { name: 'Position', exact: true }).selectOption('raised');
			await dialog.getByRole('spinbutton', { name: 'Position by (points)', exact: true }).fill('2');
			await dialog.getByRole('checkbox', { name: 'Kerning for fonts', exact: true }).check();
			await dialog.getByRole('spinbutton', { name: 'Points and above', exact: true }).fill('12');
			await dialog.getByRole('button', { name: 'OK', exact: true }).click();
			const tracked = await tree();
			const bytes = await editor.evaluate(async (el) =>
				Array.from(await (el as DocxEditorElement).saveBytes()),
			);
			const exported = (await loadDocx(new Uint8Array(bytes))).model;
			const runs = exported.blocks.flatMap((p) => (p.type === 'paragraph' ? p.runs : []));
			const run = runs.find(
				(item) => item.image || item.noteReference || item.break || item.text.includes('\n'),
			)!;
			expect(run).toMatchObject({
				fontSize: 18,
				color: '#c00000',
				smallCaps: true,
				characterSpacingTwips: 30,
				textScalePercent: 150,
				positionHalfPoints: 4,
				kerningHalfPoints: 24,
			});
			expect(run.formatRevision ?? run.revision).toMatchObject({ kind: 'formatChange' });
			await body.press('Control+z');
			expect(await tree()).toEqual(initial);
			await body.press('Control+Shift+z');
			expect(await tree()).toEqual(tracked);
			await editor.getByRole('tab', { name: 'Review', exact: true }).click();
			await editor.getByRole('button', { name: 'Reject all', exact: true }).click();
			const rejectedBytes = await editor.evaluate(async (el) =>
				Array.from(await (el as DocxEditorElement).saveBytes()),
			);
			expect(content((await loadDocx(new Uint8Array(rejectedBytes))).model)).toEqual(
				content((await loadDocx(await readFile(fixture))).model),
			);
			await body.press('Control+z');
			expect(await tree()).toEqual(tracked);
		});

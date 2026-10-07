import { fileURLToPath } from 'node:url';
import { test, expect } from '@playwright/test';
import JSZip from 'jszip';
import type { DocxEditorElement } from '../../viewers/docx/packages/web-component/src';
import { fileInput } from './helpers';

const fixture = fileURLToPath(
	new URL('../../src/core/docx/__fixtures__/review-preferences/preferences.docx', import.meta.url),
);
for (const framework of ['vanilla', 'react', 'vue', 'angular', 'svelte', 'solid'])
	test(`${framework}: edits native tracking preferences with atomic undo and export`, async ({
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
		await editor.getByRole('tab', { name: 'Review', exact: true }).click();
		await editor.getByRole('button', { name: 'Tracking options', exact: true }).click();
		const modal = editor.getByRole('dialog', { name: 'Tracking options', exact: true });
		const dialog = editor.locator('office-ui-dialog[data-title="Tracking options"]');
		await expect(modal).toBeVisible();
		await expect(
			dialog.getByRole('checkbox', { name: 'Track formatting', exact: true }),
		).not.toBeChecked();
		await expect(
			dialog.getByRole('checkbox', { name: 'Track moves', exact: true }),
		).not.toBeChecked();
		await dialog.getByRole('checkbox', { name: 'Track formatting', exact: true }).check();
		await dialog.getByRole('checkbox', { name: 'Track moves', exact: true }).check();
		await dialog.getByRole('button', { name: 'OK', exact: true }).click();
		await expect(modal).not.toBeVisible();
		const preferences = () =>
			editor.evaluate((element) => {
				const model = (element as DocxEditorElement).documentModel!;
				return { formatting: model.trackFormatting, moves: model.trackMoves };
			});
		await expect.poll(preferences).toEqual({ formatting: true, moves: true });
		await body.click();
		await page.keyboard.press('Control+z');
		await expect.poll(preferences).toEqual({ formatting: false, moves: false });
		await page.keyboard.press('Control+Shift+z');
		await expect.poll(preferences).toEqual({ formatting: true, moves: true });
		const bytes = await editor.evaluate(async (element) =>
			Array.from(await (element as DocxEditorElement).saveBytes()),
		);
		const settings = await (
			await JSZip.loadAsync(new Uint8Array(bytes))
		)
			.file('word/settings.xml')!
			.async('string');
		expect(settings).not.toContain('doNotTrackFormatting');
		expect(settings).not.toContain('doNotTrackMoves');
		expect(errors).toEqual([]);
	});

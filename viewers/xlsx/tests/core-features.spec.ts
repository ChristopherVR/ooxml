import { expect, test } from '@playwright/test';
import { resolve } from 'node:path';
import {
	editor,
	editorProperty,
	grid,
	landingFileInput,
	newWorkbook,
	openLanding,
} from './helpers';

test('Excel encrypted workbook prompts, retries and opens', async ({ page }) => {
	await newWorkbook(page);
	// Use the public load API so the current workbook remains visible behind the password prompt.
	const bytes = await import('node:fs').then((fs) => [
		...fs.readFileSync(resolve('tests/support/excel-encrypted.xlsx')),
	]);
	const loading = editor(page).evaluate(async (node, data) => {
		await (node as unknown as { load(bytes: Uint8Array, name: string): Promise<void> }).load(
			new Uint8Array(data),
			'protected.xlsx',
		);
	}, bytes);
	const prompt = editor(page).locator('[data-dialog="open-password"]');
	await expect(prompt).toBeVisible();
	await prompt.getByLabel('Password:', { exact: true }).fill('wrong');
	await prompt.getByRole('button', { name: 'OK', exact: true }).click();
	await expect(editor(page)).toContainText('The password is incorrect.');
	await expect(prompt).toBeVisible();
	await prompt.getByLabel('Password:', { exact: true }).fill('open sesame');
	await prompt.getByRole('button', { name: 'OK', exact: true }).click();
	await loading;
	await expect(prompt).toHaveCount(0);
	expect(await editorProperty<string>(page, 'fileName')).toBe('protected.xlsx');
});

test('Excel SmartArt uses the shared renderer and declares display-only support', async ({
	page,
}) => {
	await openLanding(page);
	await (await landingFileInput(page)).setInputFiles(resolve('tests/support/excel-smartart.xlsx'));
	await expect(grid(page)).toContainText('SmartArt below');
	const drawing = editor(page).locator('office-ui-smartart').first();
	await expect(drawing.locator('svg')).toBeVisible();
	await expect(drawing).toContainText('Plan');
	await expect(drawing).toContainText('Build');
	await expect(drawing).toContainText('Ship');
	await expect(editor(page).locator('.xg-smartart-notice').first()).toHaveText(
		'SmartArt (display only)',
	);
});

test('Info edits properties with undo and encrypts a downloaded workbook', async ({ page }) => {
	await newWorkbook(page);
	await editor(page).getByRole('button', { name: 'File', exact: true }).click();
	const title = editor(page).getByLabel('Title', { exact: true });
	await title.fill('Quarterly report');
	await title.blur();
	await expect
		.poll(() =>
			editorProperty<{ properties: { title?: string } }>(page, 'workbook').then(
				(w) => w.properties.title,
			),
		)
		.toBe('Quarterly report');
	await editor(page).evaluate((node) => (node as unknown as { undo(): void }).undo());
	await expect
		.poll(() =>
			editorProperty<{ properties: { title?: string } }>(page, 'workbook').then(
				(w) => w.properties.title,
			),
		)
		.toBeUndefined();
	await editor(page).getByLabel('Company', { exact: true }).fill('Example company');
	await editor(page).getByLabel('Company', { exact: true }).blur();
	await editor(page).getByLabel('Property name', { exact: true }).fill('Project');
	await editor(page).getByLabel('Property value', { exact: true }).fill('Viewer');
	await editor(page).getByRole('button', { name: 'Add text property', exact: true }).click();
	await expect(editor(page).getByLabel('Project', { exact: true })).toHaveValue('Viewer');
	await editor(page).getByRole('button', { name: 'Encrypt Workbook', exact: true }).click();
	const prompt = editor(page).locator('[data-dialog="encrypt-workbook"]');
	await prompt.getByLabel('Password:', { exact: true }).fill('secret');
	await prompt.getByLabel('Confirm password:', { exact: true }).fill('secret');
	await prompt.getByRole('button', { name: 'OK', exact: true }).click();
	await expect(
		editor(page).getByRole('button', { name: 'Remove password', exact: true }),
	).toBeVisible();
	const download = page.waitForEvent('download');
	await editor(page).evaluate((node) =>
		(node as unknown as { download(name: string): Promise<void> }).download('encrypted.xlsx'),
	);
	const path = await (await download).path();
	const { readFile } = await import('node:fs/promises');
	const { loadWorkbook } = await import('../packages/core/src/load');
	const bytes = new Uint8Array(await readFile(path!));
	await expect(loadWorkbook(bytes)).rejects.toMatchObject({ code: 'password-required' });
	const saved = await loadWorkbook(bytes, { password: 'secret' });
	expect(saved.properties.company).toBe('Example company');
	expect(saved.properties.custom).toContainEqual(
		expect.objectContaining({ name: 'Project', type: 'lpwstr', value: 'Viewer' }),
	);
});

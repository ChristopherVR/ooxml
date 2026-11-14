import JSZip from 'jszip';
import { expect, test } from '@playwright/test';
import type { DocxEditorElement } from 'docx-web-component';
import { dialogByHeading, openSample, reopenDocument } from './helpers';
import { inspectFixture } from '../../src/core/docx/test-support/inspect-fixture';

test('File > Info > Check for Issues inspects the document and removes each category, undoably', async ({
	page,
}) => {
	await openSample(page);
	await reopenDocument(page, 'inspect.docx', await inspectFixture());
	const host = page.locator('docx-editor');
	await host.locator('.dve-ribbon .file').click();
	await host.locator('.dve-backstage .item', { hasText: /^Info$/ }).click();
	await host.getByRole('button', { name: 'Check for Issues' }).click();
	const dialog = dialogByHeading(host, 'Document Inspector');
	await expect(dialog).toBeVisible();
	await expect(dialog.getByRole('checkbox')).toHaveCount(5);
	await dialog.getByRole('button', { name: 'Inspect', exact: true }).click();
	const row = (category: string) =>
		dialog.locator(`.dve-inspector-result[data-category="${category}"]`);
	for (const category of ['comments', 'properties', 'customXml', 'headersFooters', 'hiddenText'])
		await expect(row(category)).toHaveAttribute('data-state', 'found');
	await expect(row('properties')).toContainText('Reviewer names');
	await row('hiddenText').getByRole('button', { name: 'Remove All' }).click();
	await row('properties').getByRole('button', { name: 'Remove All' }).click();
	await expect(row('hiddenText')).toHaveAttribute('data-state', 'clear');
	await expect(row('properties')).toContainText('All items were removed.');
	await dialog.locator('.dve-dialog-actions').getByRole('button', { name: 'Close' }).click();
	await expect(host.locator('.ProseMirror').first()).not.toContainText('hidden words');

	const saved = async () => {
		const bytes = await host.evaluate(async (element) => {
			const blob = await (element as DocxEditorElement).save();
			return Array.from(new Uint8Array(await blob.arrayBuffer()));
		});
		return JSZip.loadAsync(Uint8Array.from(bytes));
	};
	let zip = await saved();
	const settings = (await zip.file('word/settings.xml')?.async('string')) ?? '';
	expect(settings).toContain('<w:removePersonalInformation/>');
	expect(await zip.file('word/comments.xml')?.async('string')).toContain('w:author="Author"');
	expect(zip.file('word/people.xml')).toBeNull();
	expect(await zip.file('word/document.xml')?.async('string')).not.toContain('hidden words');

	// Undo brings back the properties, then the hidden text.
	await host.locator('.ProseMirror').first().click();
	await page.keyboard.press('ControlOrMeta+z');
	await page.keyboard.press('ControlOrMeta+z');
	await expect(host.locator('.ProseMirror').first()).toContainText('hidden words');
	zip = await saved();
	expect(await zip.file('word/settings.xml')?.async('string')).not.toContain(
		'removePersonalInformation',
	);
	expect(zip.file('word/people.xml')).not.toBeNull();
});

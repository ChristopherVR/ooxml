import JSZip from 'jszip';
import { expect, test, type Page } from '@playwright/test';
import type { DocxEditorElement } from 'docx-web-component';
import { dialogByHeading, openSample, reopenDocument } from './helpers';
import { inspectFixture } from '../../src/core/docx/test-support/inspect-fixture';

/** Opens the inspector on the fixture and runs Inspect. */
async function inspect(page: Page) {
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
	const saved = async () => {
		const bytes = await host.evaluate(async (element) => {
			const blob = await (element as DocxEditorElement).save();
			return Array.from(new Uint8Array(await blob.arrayBuffer()));
		});
		return JSZip.loadAsync(Uint8Array.from(bytes));
	};
	const close = () =>
		dialog.locator('.dve-dialog-actions').getByRole('button', { name: 'Close' }).click();
	const undo = async (times: number) => {
		await host.locator('.ProseMirror').first().click();
		for (let i = 0; i < times; i++) await page.keyboard.press('ControlOrMeta+z');
	};
	return { host, row, saved, close, undo };
}

test('File > Info > Check for Issues removes hidden text and personal information, undoably', async ({
	page,
}) => {
	const { host, row, saved, close, undo } = await inspect(page);
	for (const category of ['comments', 'properties', 'customXml', 'headersFooters', 'hiddenText'])
		await expect(row(category)).toHaveAttribute('data-state', 'found');
	await expect(row('properties')).toContainText('Reviewer names');
	await row('hiddenText').getByRole('button', { name: 'Remove All' }).click();
	await row('properties').getByRole('button', { name: 'Remove All' }).click();
	await expect(row('hiddenText')).toHaveAttribute('data-state', 'clear');
	await expect(row('properties')).toContainText('All items were removed.');
	await close();
	await expect(host.locator('.ProseMirror').first()).not.toContainText('hidden words');
	let zip = await saved();
	const settings = (await zip.file('word/settings.xml')?.async('string')) ?? '';
	expect(settings).toContain('<w:removePersonalInformation/>');
	expect(await zip.file('word/comments.xml')?.async('string')).toContain('w:author="Author"');
	expect(zip.file('word/people.xml')).toBeNull();
	expect(await zip.file('word/document.xml')?.async('string')).not.toContain('hidden words');

	// Undo brings back the properties, then the hidden text.
	await undo(2);
	await expect(host.locator('.ProseMirror').first()).toContainText('hidden words');
	zip = await saved();
	expect(await zip.file('word/settings.xml')?.async('string')).not.toContain(
		'removePersonalInformation',
	);
	expect(zip.file('word/people.xml')).not.toBeNull();
});

test('Remove All clears comments and revisions, headers and footers, and custom XML, undoably', async ({
	page,
}) => {
	const { host, row, saved, close, undo } = await inspect(page);
	for (const category of ['comments', 'headersFooters', 'customXml']) {
		await row(category).getByRole('button', { name: 'Remove All' }).click();
		await expect(row(category)).toHaveAttribute('data-state', 'clear');
	}
	await close();
	let zip = await saved();
	const document = (await zip.file('word/document.xml')?.async('string')) ?? '';
	expect(document).not.toMatch(/commentReference|<w:ins /);
	expect(document).toContain('Tracked insertion.');
	expect(await zip.file('word/comments.xml')?.async('string')).not.toContain('<w:comment ');
	expect(await zip.file('word/header1.xml')?.async('string')).not.toContain('Header text');
	expect(await zip.file('word/footer1.xml')?.async('string')).not.toContain('Footer text');
	expect(zip.file('customXml/item1.xml')).toBeNull();

	await undo(3);
	zip = await saved();
	expect(await zip.file('word/document.xml')?.async('string')).toMatch(/commentReference/);
	expect(await zip.file('word/document.xml')?.async('string')).toMatch(/<w:ins /);
	expect(await zip.file('word/comments.xml')?.async('string')).toContain('A comment');
	expect(await zip.file('word/header1.xml')?.async('string')).toContain('Header text');
	expect(zip.file('customXml/item1.xml')).not.toBeNull();
});

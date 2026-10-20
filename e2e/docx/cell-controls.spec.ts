import { expect, test } from '@playwright/test';
import { savedDocument } from './saved-document';
import { fileInput } from './helpers';
import { cellControlsFixture } from './support/cell-controls-fixture';

/** The `w:sdt` of the table row that starts with `label`, as saved. */
const rowXml = (xml: string, label: string) => {
	const start = xml.indexOf(`>${label}<`);
	return xml.slice(start, xml.indexOf('</w:tr>', start));
};

for (const framework of ['vanilla', 'react', 'vue', 'angular', 'svelte', 'solid']) {
	test(`${framework}: drop-down and check box controls around table cells are operated and saved`, async ({
		page,
	}) => {
		await page.goto(`/?framework=${framework}`);
		await (
			await fileInput(page)
		).setInputFiles({
			name: 'sources.docx',
			mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
			buffer: Buffer.from(await cellControlsFixture()),
		});
		const editor = page.locator('docx-editor');
		const cells = editor.locator('.ProseMirror td[data-sdt-cell]');
		await expect(cells).toHaveCount(3);
		const temporary = cells.nth(0);
		const kept = cells.nth(1);
		await expect(kept).toHaveAttribute('data-sdt-placeholder', '1');
		await expect(kept).toHaveText('Pick a kind.');

		// A temporary control goes with its first pick and leaves the picked text in the cell.
		await temporary.click();
		await editor.getByRole('option', { name: 'Secondary', exact: true }).click();
		const first = editor.locator('.ProseMirror td').nth(1);
		await expect(first).toHaveText('Secondary');
		await expect(first).not.toHaveAttribute('data-sdt-cell');
		await expect(cells).toHaveCount(2);

		await cells.nth(0).click();
		await editor.getByRole('option', { name: 'Primary', exact: true }).click();
		await expect(cells.nth(0)).toHaveText('Primary');
		await expect(cells.nth(0)).not.toHaveAttribute('data-sdt-placeholder');
		await cells.nth(1).click();
		await expect(cells.nth(1)).toHaveText('☒');

		const { xml } = await savedDocument(page, `${framework}-cell-controls.docx`);
		const firstRow = rowXml(xml, 'First source');
		expect(firstRow).toContain('<w:t>Secondary</w:t>');
		expect(firstRow).not.toContain('<w:sdt');
		const secondRow = rowXml(xml, 'Second source');
		expect(secondRow).toContain('<w:alias w:val="Kind 82"/>');
		expect(secondRow).not.toContain('showingPlcHdr');
		expect(secondRow).toContain('<w:t>Primary</w:t>');
		expect(rowXml(xml, 'Approved')).toContain('<w14:checked w14:val="1"/>');
	});
}

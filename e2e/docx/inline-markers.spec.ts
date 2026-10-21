import { expect, test } from '@playwright/test';
import { savedDocument } from './saved-document';
import { fileInput } from './helpers';
import { inlineMarkersFixture } from './support/inline-markers-fixture';

/** Asserts that every needle occurs in `xml`, in this order. */
function expectInOrder(xml: string, needles: string[]) {
	let from = 0;
	for (const needle of needles) {
		const index = xml.indexOf(needle, from);
		expect(index, `${needle} after offset ${from}`).toBeGreaterThanOrEqual(0);
		from = index + needle.length;
	}
}

for (const framework of ['vanilla', 'react', 'vue', 'angular', 'svelte', 'solid']) {
	test(`${framework}: paragraphs with proofing marks, smart tags and symbols are edited and saved`, async ({
		page,
	}) => {
		await page.goto(`/?framework=${framework}`);
		await (
			await fileInput(page)
		).setInputFiles({
			name: 'inline-markers.docx',
			mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
			buffer: Buffer.from(await inlineMarkersFixture()),
		});
		const paragraphs = page.locator('docx-editor .ProseMirror > p');
		await expect(paragraphs.nth(0)).toHaveText('A speling mistake');
		await expect(paragraphs.nth(1)).toHaveText('Meet in Paris');
		for (const index of [0, 1, 2]) {
			await paragraphs.nth(index).click();
			await page.keyboard.press('End');
			await page.keyboard.type('!');
		}
		await expect(paragraphs.nth(0)).toHaveText('A speling mistake!');
		// The symbol, and only the symbol, is drawn in its own font.
		await expect(paragraphs.nth(2).locator('span[style*="Wingdings"]')).toHaveText('\uF0FE');

		const { xml } = await savedDocument(page, `${framework}-inline-markers.docx`);
		// The marks keep their characters although the editor merges runs of the same formatting;
		// ranges that end with their paragraph take in the text typed at its end, and the smart tag
		// stays around the text it marks.
		expectInOrder(xml, [
			'<w:proofErr w:type="spellStart"/>',
			'>speling<',
			'<w:proofErr w:type="spellEnd"/>',
			'<w:permStart w:id="1" w:edGrp="everyone"/>',
			'> mistake!<',
			'<w:permEnd w:id="1"/>',
			'<w:bookmarkStart w:id="2" w:name="Place"/>',
			'>Meet in <',
			'<w:smartTag w:uri="urn:schemas-microsoft-com:office:smarttags" w:element="City">',
			'>Paris<',
			'</w:smartTag>',
			'>!<',
			'<w:bookmarkEnd w:id="2"/>',
			'<w:ptab w:relativeTo="margin" w:alignment="right" w:leader="none"/>',
			'<w:sym w:font="Wingdings" w:char="F0FE"/>',
		]);

		// Print Layout draws the symbol in its font and the text typed after it in the run's, and
		// the positional tab sends both to the right margin.
		const editor = page.locator('docx-editor');
		await editor.getByRole('tab', { name: 'View', exact: true }).click();
		await editor.getByRole('combobox', { name: 'Layout view', exact: true }).selectOption('print');
		const line = editor.locator('.dve-print-line[data-block-id="p2"]');
		await expect(line).toBeVisible();
		const typed = line.locator('span', { hasText: '!' });
		await expect(line.locator('span', { hasText: '\uF0FE' })).toHaveCSS('font-family', /Wingdings/);
		await expect(typed).not.toHaveCSS('font-family', /Wingdings/);
		const lineBox = (await line.boundingBox())!;
		const typedBox = (await typed.boundingBox())!;
		expect(Math.abs(typedBox.x + typedBox.width - (lineBox.x + lineBox.width))).toBeLessThan(6);
	});
}

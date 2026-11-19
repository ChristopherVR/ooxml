/**
 * Accessibility survives an edited save: open each fixture, change it through the UI, download the
 * file the suite writes and compare its alt text, headings, table headers and language with the
 * original package.
 */
import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';
import type { Page } from '@playwright/test';
import { EXCEL, POWERPOINT, WORD } from './support/fixtures';
import { expect, FILES, openSurface, test, type Product } from './support/suite';

test.use({ viewport: { width: 1440, height: 900 } });

type Parts = Record<string, string>;
const parts = async (bytes: Uint8Array | Buffer): Promise<Parts> => {
	const zip = await JSZip.loadAsync(bytes);
	const out: Parts = {};
	for (const [name, file] of Object.entries(zip.files))
		if (/\.(xml|rels)$/u.test(name)) out[name] = await file.async('string');
	return out;
};
const count = (text: string | undefined, pattern: RegExp) =>
	(text ?? '').match(pattern)?.length ?? 0;
const joined = (files: Parts, name: RegExp) =>
	Object.entries(files)
		.filter(([path]) => name.test(path))
		.map(([, text]) => text)
		.join('\n');

/** The accessibility facts of a saved package. */
const FACTS: Partial<Record<Product, (files: Parts) => Record<string, number>>> = {
	docx: (f) => {
		const body = f['word/document.xml'];
		return {
			'Title paragraphs': count(body, /w:pStyle w:val="Title"/gu),
			'Heading 1 paragraphs': count(body, /w:pStyle w:val="Heading1"/gu),
			'Heading 2 paragraphs': count(body, /w:pStyle w:val="Heading2"/gu),
			'heading 1 outline level': count(
				f['word/styles.xml'],
				/w:styleId="Heading1"[\s\S]*?<w:outlineLvl w:val="0"/gu,
			),
			'picture alt text': count(body, new RegExp(`descr="${WORD.alt}"`, 'gu')),
			'repeated header row': count(body, /<w:tblHeader/gu),
			'table caption': count(body, new RegExp(`<w:tblCaption w:val="${WORD.caption}"`, 'gu')),
			'table description': count(body, /<w:tblDescription /gu),
			'French run': count(body, /<w:lang w:val="fr-FR"/gu),
			'Arabic run': count(body, /<w:lang[^>]*w:bidi="ar-SA"/gu),
			'right-to-left paragraph': count(body, /<w:bidi\/>/gu),
			'default language': count(
				f['word/styles.xml'],
				new RegExp(`<w:lang w:val="${WORD.language}"`, 'gu'),
			),
			'document language': count(
				f['docProps/core.xml'],
				new RegExp(`<dc:language>${WORD.language}<`, 'gu'),
			),
		};
	},
	xlsx: (f) => {
		const tables = joined(f, /^xl\/tables\/table\d+\.xml$/u);
		return {
			'table header row':
				count(tables, /headerRowCount="1"/gu) +
				count(tables, /<table(?![^>]*headerRowCount)[^>]*>/gu),
			'table alt text': count(tables, new RegExp(`altText="${EXCEL.tableAlt}"`, 'gu')),
			'table alt summary': count(
				tables,
				new RegExp(`altTextSummary="${EXCEL.tableSummary}"`, 'gu'),
			),
			'picture alt text': count(
				joined(f, /^xl\/drawings\/drawing\d+\.xml$/u),
				new RegExp(`descr="${EXCEL.pictureAlt}"`, 'gu'),
			),
		};
	},
	pptx: (f) => {
		const slide = f['ppt/slides/slide1.xml'];
		return {
			'picture alt text': count(slide, new RegExp(`descr="${POWERPOINT.alt}"`, 'gu')),
			'decorative picture': count(slide, /<adec:decorative [^>]*val="1"/gu),
			'shape order': count(
				slide,
				/<p:cNvPr id="2" name="Described Picture"[\s\S]*<p:cNvPr id="3" name="Decorative Picture"/gu,
			),
		};
	},
};

/** One edit made the way a user would, so the save really rewrites the package. */
const EDIT: Partial<Record<Product, (page: Page) => Promise<void>>> = {
	docx: async (page) => {
		await page.locator('docx-editor .ProseMirror').first().focus();
		await page.keyboard.press('Control+End');
		await page.keyboard.type(' Edited.');
	},
	xlsx: async (page) => {
		await page
			.locator('xlsx-editor')
			.evaluate((editor) => (editor as HTMLElement & { focusGrid(): void }).focusGrid());
		await page.keyboard.press('Control+End');
		await page.keyboard.press('ArrowDown');
		await page.keyboard.type('9');
		await page.keyboard.press('Enter');
	},
	pptx: async (page) => {
		await page.getByRole('button', { name: 'Add Slide' }).first().click();
	},
};

/** Whether the saved package carries the edit, so the check is not reading the original back. */
const EDITED: Partial<Record<Product, (files: Parts) => boolean>> = {
	docx: (f) => (f['word/document.xml'] ?? '').includes('Edited.</w:t>'),
	xlsx: (f) => (f['xl/worksheets/sheet1.xml'] ?? '').includes('<v>9</v>'),
	pptx: (f) => Object.keys(f).some((name) => name === 'ppt/slides/slide2.xml'),
};

const unsaved = (page: Page) =>
	page.evaluate(() =>
		Boolean(
			document
				.querySelector('#tabs [role="tab"][aria-selected="true"]')
				?.textContent?.includes('•'),
		),
	);

for (const product of Object.keys(FACTS) as Product[]) {
	const facts = FACTS[product]!;
	const edit = EDIT[product]!;
	test(`${product}: an edited save keeps alt text, headings, table headers and language`, async ({
		page,
	}) => {
		const original = facts(await parts((await FILES[product]()).buffer));
		for (const [fact, value] of Object.entries(original))
			expect(value, `the fixture has ${fact}`).toBeGreaterThan(0);
		await openSurface(page, product);
		await edit(page);
		await expect.poll(() => unsaved(page), 'the edit marks the document unsaved').toBe(true);
		await page.waitForTimeout(1000);
		await page.keyboard.press('Control+s');
		await expect.poll(() => unsaved(page), 'Ctrl+S saves it').toBe(false);
		await page.locator('#actions-more').click();
		const [download] = await Promise.all([
			page.waitForEvent('download'),
			page
				.locator('#actions-menu [role="menuitem"]')
				.filter({ hasText: /^Download$/u })
				.first()
				.click(),
		]);
		const saved = await parts(await readFile(await download.path()));
		expect(EDITED[product]!(saved), 'the download holds the edit').toBe(true);
		expect(facts(saved)).toEqual(original);
	});
}

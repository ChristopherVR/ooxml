// Worksheet parts with comments or CDATA load on the fast cell reader with exactly the model the
// DOM path builds. The DOM path is forced with a processing instruction in the cell data, which the
// fast reader still declines and no reader looks at.
import { readFileSync } from 'node:fs';
import path from 'node:path';
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { LiteDeclinedError } from '../../xml/index';
import { loadXlsx } from './load';
import { splitSheetData } from './sheet-data-source';

const fixture = new Uint8Array(
	readFileSync(path.join(import.meta.dirname, '..', '__fixtures__', 'excel-features.xlsx')),
);

const XML_ESCAPES: Record<string, string> = {
	'&amp;': '&',
	'&lt;': '<',
	'&gt;': '>',
	'&quot;': '"',
	'&apos;': "'",
};

/** Comments before the root, in the header and at every level of the cell data. */
const withComments = (xml: string): string =>
	xml
		.replace(/^(<\?xml[^>]*\?>)/, '$1<!-- written by another tool -->')
		.replace(/<sheetData/, '<!-- <sheetData> --><sheetData')
		.replace(/<\/row>/g, '</row><!-- end of a row -->')
		.replace(/<c /g, '<!-- a cell --><c ')
		.replace(/<v>/g, '<v><!-- a value -->')
		.replace(/<\/sheetData>/, '</sheetData><!-- </sheetData> -->');

/** Every formula's text as a CDATA section. */
const withCdata = (xml: string): string =>
	xml.replace(
		/<f([^>]*)>([^<]+)<\/f>/g,
		(_, attrs: string, text: string) =>
			`<f${attrs}><![CDATA[${text.replace(/&(?:amp|lt|gt|quot|apos);/g, (e) => XML_ESCAPES[e] ?? e)}]]></f>`,
	);

/** The same part, made to take the DOM path (not the `<sheetData>` quoted in a comment). */
const forceDom = (xml: string): string => xml.replace(/(<sheetData[^>]*>)(?!\s*-->)/, '$1<?keep?>');

async function loadEdited(edit: (xml: string) => string) {
	const zip = await JSZip.loadAsync(fixture);
	const parts = Object.keys(zip.files).filter((name) => /^xl\/worksheets\/[^/]+\.xml$/.test(name));
	const edited: string[] = [];
	let changed = 0;
	for (const name of parts) {
		const original = await zip.file(name)!.async('string');
		const xml = edit(original);
		if (xml !== original) changed++;
		edited.push(xml);
		zip.file(name, xml);
	}
	const workbook = await loadXlsx(await zip.generateAsync({ type: 'uint8array' }));
	return { workbook, edited, changed };
}

describe('worksheets with comments or CDATA', () => {
	it.each([
		['comments', withComments],
		['CDATA formulas', withCdata],
	])('load %s on the fast reader exactly as the DOM path does', async (_, edit) => {
		const fast = await loadEdited(edit);
		const dom = await loadEdited((xml) => forceDom(edit(xml)));
		expect(fast.changed).toBeGreaterThan(0);
		for (const xml of fast.edited) {
			expect(xml).not.toBe(forceDom(xml));
			expect(() => [...splitSheetData(xml)!.rows]).not.toThrow();
		}
		for (const xml of dom.edited)
			expect(() => [...splitSheetData(xml)!.rows]).toThrow(LiteDeclinedError);
		expect(fast.workbook.sheets).toStrictEqual(dom.workbook.sheets);
		expect(fast.workbook.warnings).toStrictEqual(dom.workbook.warnings);
		const plain = await loadXlsx(fixture);
		expect(fast.workbook.sheets.map((sheet) => sheet.rows)).toStrictEqual(
			plain.sheets.map((sheet) => sheet.rows),
		);
	});
});

describe('a worksheet the light reader declines part way through', () => {
	it('is read again whole, into the model and warnings of the DOM path', async () => {
		// A missing shared string in the first rows (a warning) and markup the light reader leaves
		// to the XML parser after the last row, so the rows are read before it gives up.
		const late = (xml: string): string =>
			xml
				.replace(/(<c [^>]*t="s"[^>]*><v>)\d+(<\/v>)/, '$199999$2')
				.replace('</sheetData>', '<?late?></sheetData>');
		const streamed = await loadEdited(late);
		const dom = await loadEdited((xml) => forceDom(late(xml)));
		expect(streamed.changed).toBeGreaterThan(0);
		for (const xml of streamed.edited) {
			const rows = splitSheetData(xml)!.rows[Symbol.iterator]();
			expect(rows.next().done).toBe(false);
			expect(() => {
				while (!rows.next().done);
			}).toThrow(LiteDeclinedError);
		}
		expect(streamed.workbook.sheets).toStrictEqual(dom.workbook.sheets);
		expect(streamed.workbook.warnings).toStrictEqual(dom.workbook.warnings);
		expect(streamed.workbook.warnings).toContain('A cell refers to a missing shared string.');
	});
});

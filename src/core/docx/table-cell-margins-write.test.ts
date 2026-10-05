import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { loadDocx } from './parse.js';
import type { Table } from './model.js';
import { signedTwips } from './units.js';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
async function fixture(
	margins = '<w:top w:w="30" w:type="dxa" ext:keep="top"/><w:left w:w="50" w:type="dxa"/><w:start w:w="100" w:type="dxa" ext:keep="start"/><w:bottom w:w="20" w:type="dxa"/><w:end w:w="80" w:type="dxa"/>',
) {
	const zip = new JSZip();
	zip.file(
		'word/document.xml',
		`<w:document xmlns:w="${W}" xmlns:ext="urn:test"><w:body><w:tbl><w:tblPr><w:tblCellMar><w:left w:w="200" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/><w:tcBorders><w:top w:val="single"/></w:tcBorders><w:shd w:fill="EEEEEE"/><w:tcMar>${margins}</w:tcMar><ext:keep ext:value="yes"/></w:tcPr><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:sectPr/></w:body></w:document>`,
	);
	return zip.generateAsync({ type: 'uint8array' });
}
const tableOf = (loaded: Awaited<ReturnType<typeof loadDocx>>) => loaded.model.blocks[0] as Table;
const cellOf = (loaded: Awaited<ReturnType<typeof loadDocx>>) => tableOf(loaded).rows[0]![0]!;
async function xmlOf(bytes: Uint8Array) {
	return (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string');
}

describe('preservation-safe cell margins', () => {
	it('edits one logical side and preserves untouched margins and cell metadata', async () => {
		const loaded = await loadDocx(await fixture());
		expect(cellOf(loaded).margins?.left).toBe(100);
		cellOf(loaded).margins!.left = signedTwips(150);
		const bytes = await loaded.save();
		const xml = await xmlOf(bytes);
		expect(xml).toContain('<w:start w:w="150" w:type="dxa" ext:keep="start"/>');
		expect(xml).not.toContain('<w:left w:w="50"');
		expect(xml).toContain('<w:top w:w="30" w:type="dxa" ext:keep="top"/>');
		expect(xml).toContain('<w:end w:w="80" w:type="dxa"/>');
		expect(xml).toContain('<w:tcW w:w="2000" w:type="dxa"/>');
		expect(xml).toContain('<w:tcBorders><w:top w:val="single"/></w:tcBorders>');
		expect(xml).toContain('<w:shd w:fill="EEEEEE"/>');
		expect(xml).toContain('<ext:keep ext:value="yes"/>');
		expect(cellOf(await loadDocx(bytes)).margins?.left).toBe(150);
	});

	it('writes explicit zero, then clears both aliases to restore inheritance', async () => {
		const loaded = await loadDocx(await fixture());
		cellOf(loaded).margins!.left = signedTwips(0);
		const zero = await loadDocx(await loaded.save());
		expect(cellOf(zero).margins?.left).toBe(0);
		delete cellOf(zero).margins!.left;
		const cleared = await loadDocx(await zero.save());
		expect(cellOf(cleared).margins?.left).toBeUndefined();
		expect(tableOf(cleared).cellMargins?.left).toBe(200);
		delete cellOf(cleared).margins;
		expect(await xmlOf(await cleared.save())).not.toContain('<w:tcMar');
	});

	it('keeps no-op and neighboring text edits from rewriting imported margins', async () => {
		const bytes = await fixture();
		const loaded = await loadDocx(bytes);
		expect(await loaded.save()).toEqual(bytes);
		cellOf(loaded).paragraphs[0]!.runs[0]!.text = 'Changed';
		const before = (await xmlOf(bytes)).match(/<w:tcMar>.*?<\/w:tcMar>/)?.[0];
		expect(await xmlOf(await loaded.save())).toContain(before);
	});

	it('parses signed and nil margins, preserves unsupported widths, and rejects fractional edits', async () => {
		const loaded = await loadDocx(
			await fixture(
				'<w:top w:w="-30" w:type="dxa"/><w:left w:w="100" w:type="nil"/><w:end w:w="20" w:type="pct"/>',
			),
		);
		expect(cellOf(loaded).margins).toEqual({ top: -30, left: 0 });
		cellOf(loaded).margins!.top = signedTwips(60);
		expect(await xmlOf(await loaded.save())).toContain('<w:end w:w="20" w:type="pct"/>');
		cellOf(loaded).margins!.top = 1.5 as ReturnType<typeof signedTwips>;
		await expect(loaded.save()).rejects.toThrow('ST_SignedTwipsMeasure');
		cellOf(loaded).margins!.top = signedTwips(-60);
		await expect(loaded.save()).rejects.toThrow('nonnegative whole twips');
	});
});

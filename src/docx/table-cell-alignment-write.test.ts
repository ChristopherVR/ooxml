import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { loadDocx } from './parse.js';
import type { Table, TableCell } from './model.js';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
async function fixture(align: TableCell['verticalAlign'] = 'bottom') {
	const zip = new JSZip();
	zip.file(
		'word/document.xml',
		`<w:document xmlns:w="${W}" xmlns:ext="urn:test"><w:body><w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:tcPr ext:keep="cell"><w:tcW w:w="2000" w:type="dxa"/><w:gridSpan w:val="2"/><w:tcBorders><w:top w:val="single"/></w:tcBorders><w:shd w:fill="EEEEEE"/><w:tcMar><w:left w:w="100" w:type="dxa"/></w:tcMar>${align ? `<w:vAlign w:val="${align}" ext:keep="alignment"/>` : ''}<w:hideMark/><ext:keep ext:value="yes"/></w:tcPr><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:sectPr/></w:body></w:document>`,
	);
	return zip.generateAsync({ type: 'uint8array' });
}
const cellOf = (loaded: Awaited<ReturnType<typeof loadDocx>>) =>
	(loaded.model.blocks[0] as Table).rows[0]![0]!;
async function xmlOf(bytes: Uint8Array) {
	return (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string');
}

describe('preservation-safe cell vertical alignment', () => {
	it.each(['top', 'center', 'bottom'] as const)(
		'roundtrips explicit %s while preserving merged cell metadata',
		async (value) => {
			const loaded = await loadDocx(await fixture(value === 'bottom' ? 'top' : 'bottom'));
			cellOf(loaded).verticalAlign = value;
			const bytes = await loaded.save();
			const xml = await xmlOf(bytes);
			expect(xml).toContain(`<w:vAlign w:val="${value}" ext:keep="alignment"/>`);
			expect(xml).toContain('<w:tcPr ext:keep="cell">');
			expect(xml).toContain('<w:tcW w:w="2000" w:type="dxa"/>');
			expect(xml).toContain('<w:gridSpan w:val="2"/>');
			expect(xml).toContain('<w:tcBorders><w:top w:val="single"/></w:tcBorders>');
			expect(xml).toContain('<w:shd w:fill="EEEEEE"/>');
			expect(xml).toContain('<w:tcMar><w:left w:w="100" w:type="dxa"/></w:tcMar>');
			expect(xml).toContain('<ext:keep ext:value="yes"/>');
			expect(xml.indexOf('<w:tcMar')).toBeLessThan(xml.indexOf('<w:vAlign'));
			expect(xml.indexOf('<w:vAlign')).toBeLessThan(xml.indexOf('<w:hideMark'));
			expect(cellOf(await loadDocx(bytes)).verticalAlign).toBe(value);
		},
	);

	it('removes only the override when restoring default inheritance', async () => {
		const loaded = await loadDocx(await fixture());
		delete cellOf(loaded).verticalAlign;
		const bytes = await loaded.save();
		expect(await xmlOf(bytes)).not.toContain('<w:vAlign');
		expect(await xmlOf(bytes)).toContain('<w:tcPr ext:keep="cell">');
		expect(await xmlOf(bytes)).toContain('<ext:keep ext:value="yes"/>');
		expect(cellOf(await loadDocx(bytes)).verticalAlign).toBeUndefined();
	});

	it('preserves imported both alignment on no-op and neighboring text edits', async () => {
		const bytes = await fixture('both');
		const loaded = await loadDocx(bytes);
		expect(await loaded.save()).toEqual(bytes);
		cellOf(loaded).paragraphs[0]!.runs[0]!.text = 'Edited';
		expect(await xmlOf(await loaded.save())).toContain(
			'<w:vAlign w:val="both" ext:keep="alignment"/>',
		);
	});

	it('rejects invalid alignment tokens instead of writing invalid OOXML', async () => {
		const loaded = await loadDocx(await fixture());
		cellOf(loaded).verticalAlign = 'middle' as NonNullable<TableCell['verticalAlign']>;
		await expect(loaded.save()).rejects.toThrow('ST_VerticalJc');
	});
});

import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { loadDocx } from './parse.js';
import type { Table } from './model.js';
import { resolveCellVisuals } from './table-visuals.js';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
async function fixture() {
	const zip = new JSZip();
	zip.file(
		'word/document.xml',
		`<w:document xmlns:w="${W}" xmlns:ext="urn:test"><w:body><w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/><w:tcBorders><w:top w:val="single"/></w:tcBorders><w:shd w:val="pct20" w:fill="EEEEEE" w:themeFill="accent1" w:themeFillTint="80"/><w:tcMar><w:left w:w="100" w:type="dxa"/></w:tcMar><ext:keep ext:value="yes"/></w:tcPr><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:sectPr/></w:body></w:document>`,
	);
	return zip.generateAsync({ type: 'uint8array' });
}

describe('preservation-safe cell shading', () => {
	it('changes the fill, clears theme overrides and preserves all unrelated cell properties', async () => {
		const loaded = await loadDocx(await fixture());
		const cell = (loaded.model.blocks[0] as Table).rows[0]![0]!;
		cell.shadingFill = '#123456';
		delete cell.shadingThemeFill;
		const bytes = await loaded.save();
		const xml = await (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string');
		expect(xml).toContain('w:fill="123456"');
		expect(xml).not.toContain('themeFill');
		expect(xml).not.toContain('pct20');
		expect(xml).toContain('<w:tcW w:w="2000" w:type="dxa"/>');
		expect(xml).toContain('<w:tcBorders><w:top w:val="single"/></w:tcBorders>');
		expect(xml).toContain('<w:tcMar><w:left w:w="100" w:type="dxa"/></w:tcMar>');
		expect(xml).toContain('<ext:keep ext:value="yes"/>');
		expect(xml.indexOf('<w:shd')).toBeLessThan(xml.indexOf('<w:tcMar'));
		const reopened = await loadDocx(bytes);
		expect((reopened.model.blocks[0] as Table).rows[0]![0]!.shadingFill).toBe('#123456');
	});

	it('writes explicit No Color into a plain cell and keeps it after reopening', async () => {
		const loaded = await loadDocx(await fixture());
		const table = loaded.model.blocks[0] as Table;
		table.rows[0]![1]!.shadingFill = 'auto';
		const reopened = await loadDocx(await loaded.save());
		const next = reopened.model.blocks[0] as Table;
		expect(next.rows[0]![1]!.shadingFill).toBe('auto');
		next.style = 'Shaded';
		const visuals = resolveCellVisuals(
			next,
			next.rows[0]![1]!,
			{ row: 0, lastRow: 0, column: 1, lastColumn: 1, rowCount: 1, columnCount: 2 },
			{
				warnings: [],
				styles: { Shaded: { id: 'Shaded', conditional: {}, shadingFill: '#FF0000' } },
			},
		);
		expect(visuals.shadingFill).toBe('auto');
	});

	it('preserves untouched patterned/theme shading when editing text', async () => {
		const bytes = await fixture();
		const loaded = await loadDocx(bytes);
		expect(await loaded.save()).toEqual(bytes);
		(loaded.model.blocks[0] as Table).rows[0]![0]!.paragraphs[0]!.runs[0]!.text = 'Updated';
		const xml = await (
			await JSZip.loadAsync(await loaded.save())
		)
			.file('word/document.xml')!
			.async('string');
		expect(xml).toContain(
			'<w:shd w:val="pct20" w:fill="EEEEEE" w:themeFill="accent1" w:themeFillTint="80"/>',
		);
	});

	it('removes a direct fill when the caller requests inheritance and rejects invalid colors', async () => {
		const loaded = await loadDocx(await fixture());
		const cell = (loaded.model.blocks[0] as Table).rows[0]![0]!;
		cell.shadingFill = 'not-a-color';
		await expect(loaded.save()).rejects.toThrow('ST_HexColor');
		delete cell.shadingFill;
		delete cell.shadingThemeFill;
		const xml = await (
			await JSZip.loadAsync(await loaded.save())
		)
			.file('word/document.xml')!
			.async('string');
		expect(xml).not.toContain('<w:shd');
	});
});

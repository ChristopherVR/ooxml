import { expect, it } from 'vitest';
import JSZip from 'jszip';
import { loadDocx } from './parse';
import { parseTable } from './block-parser';
import { parseXml } from './xml';
import { parseTableStyleCatalog } from './table-styles';
import { resolveTableStyleFormatting } from './resolve-table';
import type { Table } from './model';

const ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const tableXml = (attrs: string) =>
	`<w:tbl xmlns:w="${ns}"><w:tblPr><w:tblStyle w:val="Audit"/><w:tblLook ${attrs}/></w:tblPr><w:tr><w:tc><w:p><w:r><w:t>Before</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`;
const parse = (attrs: string) => parseTable(parseXml(tableXml(attrs)).documentElement, 't0');
const catalog = parseTableStyleCatalog(
	`<w:styles xmlns:w="${ns}"><w:style w:type="table" w:styleId="Audit"><w:tblStylePr w:type="firstRow"><w:tcPr><w:shd w:fill="FF0000"/></w:tcPr></w:tblStylePr><w:tblStylePr w:type="band1Horz"><w:tcPr><w:shd w:fill="0000FF"/></w:tcPr></w:tblStylePr></w:style></w:styles>`,
);

it.each([
	['0020', 'firstRow'],
	['0040', 'lastRow'],
	['0080', 'firstColumn'],
	['0100', 'lastColumn'],
	['0200', 'noHBand'],
	['0400', 'noVBand'],
] as const)('decodes packed table-look bit %s into %s', (val, flag) => {
	const look = parse(`w:val="${val}"`).look!;
	expect(look[flag]).toBe(true);
	expect(Object.values(look).filter(Boolean)).toHaveLength(1);
	expect(Object.keys(look)).toHaveLength(6);
});

it('resolves first-row formatting and disabled banding from a packed look', () => {
	const table = parse('w:val="0660"');
	expect(table.look).toEqual({
		firstRow: true,
		lastRow: true,
		firstColumn: false,
		lastColumn: false,
		noHBand: true,
		noVBand: true,
	});
	expect(resolveTableStyleFormatting(table.style, catalog, table.look, 0, 3, 0, 2)).toEqual({
		shadingFill: '#FF0000',
	});
	expect(resolveTableStyleFormatting(table.style, catalog, table.look, 1, 3, 0, 2)).toEqual({});
	expect(parse('w:val="04a0"').look?.firstColumn).toBe(true);
	expect(Object.values(parse('w:val="0000"').look!).every((value) => !value)).toBe(true);
});

it.each(['firstRow', 'lastRow', 'firstColumn', 'lastColumn', 'noHBand', 'noVBand'] as const)(
	'ignores all packed flags when a named %s flag is present, including explicit off',
	(flag) => {
		expect(parse(`w:val="07E0" w:${flag}="0"`).look).toEqual({ [flag]: false });
	},
);

it('does not use packed flags to repair an invalid but present named flag', () => {
	expect(parse('w:val="07E0" w:firstRow="invalid"').look).toBeUndefined();
});

it.each(['bad', 'xyz0', '0x20', '00200', '-020'])(
	'does not accept malformed short hex %s',
	(val) => {
		expect(parse(`w:val="${val}"`).look).toBeUndefined();
	},
);

it('preserves packed source flags through a cell text edit, save and reparse', async () => {
	const zip = new JSZip();
	zip.file(
		'word/document.xml',
		`<w:document xmlns:w="${ns}"><w:body>${tableXml('w:val="0660"')}<w:sectPr/></w:body></w:document>`,
	);
	const loaded = await loadDocx(await zip.generateAsync({ type: 'uint8array' }));
	const table = loaded.model.blocks[0] as Table;
	table.rows[0]![0]!.paragraphs[0]!.runs[0]!.text = 'After';
	const saved = await loaded.save();
	const xml = await (await JSZip.loadAsync(saved)).file('word/document.xml')!.async('string');
	expect(xml).toContain('<w:tblLook w:val="0660"/>');
	expect(xml).toContain('<w:t>After</w:t>');
	const reparsed = (await loadDocx(saved)).model.blocks[0] as Table;
	expect(reparsed.look).toEqual(table.look);
	expect(resolveTableStyleFormatting(reparsed.style, catalog, reparsed.look, 0, 3, 0, 2)).toEqual({
		shadingFill: '#FF0000',
	});
});

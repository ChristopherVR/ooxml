// Builds the documents `table-structure-acceptance.ps1` reopens in Word. Run with bun from src/core:
//   bun scripts/docx/table-structure-acceptance-build.ts <output folder>
// Each file is a merged or nested table edited by the table structure engine and saved by the
// merge-aware writer (`loadDocx(...).save`). `manifest.json` lists, per file and per table row,
// the text of every cell (`w:tc`), which Word must report from `Cells` after opening the file.
import { mkdirSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import JSZip from 'jszip';
import { loadDocx } from '../../docx/index';
import type { Table, TableCell } from '../../docx/model';
import { deleteRow, insertColumn, insertRow } from '../../docx/table-structure';
import { mergeCells, splitCell } from '../../docx/table-structure-merge';

const ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const p = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
const tc = (text: string, props = '') =>
	`<w:tc><w:tcPr>${props}</w:tcPr>${text ? p(text) : '<w:p/>'}</w:tc>`;
const w = (n: number) => `<w:tcW w:w="${n}" w:type="dxa"/>`;
const cols = (...c: number[]) =>
	`<w:tblGrid>${c.map((n) => `<w:gridCol w:w="${n}"/>`).join('')}</w:tblGrid>`;

const MERGED =
	`<w:document xmlns:w="${ns}"><w:body><w:tbl><w:tblPr><w:tblW w:w="9000" w:type="dxa"/><w:tblBorders><w:top w:val="single" w:sz="4"/><w:left w:val="single" w:sz="4"/><w:bottom w:val="single" w:sz="4"/><w:right w:val="single" w:sz="4"/><w:insideH w:val="single" w:sz="4"/><w:insideV w:val="single" w:sz="4"/></w:tblBorders></w:tblPr>${cols(3000, 3000, 3000)}` +
	`<w:tr>${tc('A', w(6000) + '<w:gridSpan w:val="2"/>')}${tc('V', w(3000) + '<w:vMerge w:val="restart"/>')}</w:tr>` +
	`<w:tr>${tc('D', w(3000))}${tc('E', w(3000))}${tc('', w(3000) + '<w:vMerge/>')}</w:tr>` +
	`<w:tr>${tc('G', w(3000))}${tc('H', w(3000))}${tc('', w(3000) + '<w:vMerge/>')}</w:tr>` +
	`</w:tbl><w:sectPr/></w:body></w:document>`;
const PLAIN =
	`<w:document xmlns:w="${ns}"><w:body><w:tbl><w:tblPr/>${cols(3000, 3000, 3000)}` +
	[0, 1, 2]
		.map((r) => `<w:tr>${[0, 1, 2].map((c) => tc(`r${r}c${c}`, w(3000))).join('')}</w:tr>`)
		.join('') +
	`</w:tbl><w:sectPr/></w:body></w:document>`;
const NESTED =
	`<w:document xmlns:w="${ns}"><w:body><w:tbl><w:tblPr/>${cols(4500, 4500)}<w:tr>` +
	`<w:tc><w:tcPr>${w(4500)}</w:tcPr>${p('outer')}<w:tbl><w:tblPr/>${cols(2000, 2000)}<w:tr>${tc('i1', w(2000))}${tc('i2', w(2000))}</w:tr></w:tbl><w:p/></w:tc>${tc('B', w(4500))}</w:tr>` +
	`<w:tr>${tc('C', w(4500))}${tc('D', w(4500))}</w:tr></w:tbl><w:sectPr/></w:body></w:document>`;

let serial = 0;
const newId = () => `acc-${++serial}`;
const textOf = (cell: TableCell) =>
	cell.paragraphs.map((q) => q.runs.map((r) => r.text).join('')).join('\n');
const must = <T>(value: T | undefined, what: string): T => {
	if (value === undefined) throw new Error(`${what} was refused`);
	return value;
};

async function open(xml: string) {
	const zip = new JSZip();
	zip.file(
		'[Content_Types].xml',
		'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>',
	);
	zip.file(
		'_rels/.rels',
		'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
	);
	zip.file('word/document.xml', xml);
	const loaded = await loadDocx(await zip.generateAsync({ type: 'uint8array' }));
	const table = loaded.model.blocks[0];
	if (table?.type !== 'table') throw new Error('No table');
	return { loaded, table };
}

function put(table: Table, edited: Table): void {
	table.rows = edited.rows;
	if (edited.grid) table.grid = edited.grid;
	if (edited.rowProperties) table.rowProperties = edited.rowProperties;
	table.structureEditable = false;
}

const outDir = process.argv[2];
if (!outDir) throw new Error('Usage: table-structure-acceptance-build.ts <output folder>');
mkdirSync(outDir, { recursive: true });
const manifest: Record<string, string[][]> = {};

async function emit(
	name: string,
	xml: string,
	edit: (table: Table) => void,
	inner?: (table: Table) => Table,
) {
	const { loaded, table } = await open(xml);
	edit(table);
	writeFileSync(join(outDir, name), await loaded.save());
	const shown = inner ? inner(table) : table;
	// Word lists a vertical merge once, at its first row.
	manifest[name] = shown.rows.map((row) =>
		row.filter((cell) => cell.verticalMerge !== 'continue').map(textOf),
	);
}

await emit('merged-insert-row.docx', MERGED, (t) =>
	put(t, must(insertRow(t, 1, 'below', newId), 'insertRow')),
);
await emit('merged-delete-row.docx', MERGED, (t) => put(t, must(deleteRow(t, 0), 'deleteRow')));
await emit('merged-insert-column.docx', MERGED, (t) =>
	put(t, must(insertColumn(t, 2, 'right', newId), 'insertColumn')),
);
await emit('plain-merge.docx', PLAIN, (t) =>
	put(t, must(mergeCells(t, { firstRow: 0, lastRow: 1, firstCol: 0, lastCol: 1 }, newId), 'merge')),
);
await emit('plain-split.docx', PLAIN, (t) =>
	put(t, must(splitCell(t, 1, 1, 2, 2, newId), 'splitCell')),
);
const innerOf = (t: Table): Table => {
	const inner = t.rows[0]?.[0]?.sourceBlocks?.find((b): b is Table => b.type === 'table');
	if (!inner) throw new Error('No nested table');
	return inner;
};
await emit(
	'nested-insert-column.docx',
	NESTED,
	(t) => put(innerOf(t), must(insertColumn(innerOf(t), 0, 'right', newId), 'nested insertColumn')),
	innerOf,
);
// The outer table of that file is as authored: cell texts other than the one holding the nested table.
manifest['nested-insert-column.docx#outer'] = [
	['', 'B'],
	['C', 'D'],
];
writeFileSync(join(outDir, 'manifest.json'), JSON.stringify(manifest, null, 2));
console.log(`Wrote ${Object.keys(manifest).length} documents to ${outDir}`);

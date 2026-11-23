import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import {
	createDocument,
	eighthPoints,
	loadDocx,
	saveDocx,
	setTableAlignment,
	setTableBorders,
	setTableGridWidths,
	setTableIndent,
	setTableLook,
	setTableStyle,
	setTableWidth,
	signedTwips,
	twips,
	updateTableProperties,
	type Table,
} from './index';
import { expectTable } from './test-support/access';

const ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

function makeTable(): Table {
	return {
		type: 'table',
		id: 'tbl-1',
		grid: [twips(2000), twips(3000)],
		rows: [
			[
				{
					widthTwips: twips(2000),
					paragraphs: [{ type: 'paragraph', id: 'p-r0c0', runs: [{ text: 'A' }] }],
				},
				{
					widthTwips: twips(3000),
					paragraphs: [{ type: 'paragraph', id: 'p-r0c1', runs: [{ text: 'B' }] }],
				},
			],
			[
				{
					widthTwips: twips(2000),
					paragraphs: [{ type: 'paragraph', id: 'p-r1c0', runs: [{ text: 'C' }] }],
				},
				{
					widthTwips: twips(3000),
					paragraphs: [{ type: 'paragraph', id: 'p-r1c1', runs: [{ text: 'D' }] }],
				},
			],
		],
	};
}

async function zipXml(bodyXml: string): Promise<Uint8Array> {
	const zip = new JSZip();
	zip.file(
		'word/document.xml',
		`<w:document xmlns:w="${ns}"><w:body>${bodyXml}<w:sectPr/></w:body></w:document>`,
	);
	return zip.generateAsync({ type: 'uint8array' });
}

async function readDocumentXml(bytes: Uint8Array): Promise<string> {
	const zip = await JSZip.loadAsync(bytes);
	return (await zip.file('word/document.xml')?.async('string')) ?? '';
}

describe('table-properties-edit pure operations', () => {
	it('updates width between dxa, pct and auto without mutating the input table', () => {
		const base = makeTable();
		const dxa = setTableWidth(base, { type: 'dxa', value: 7200 });
		expect(base.widthTwips).toBeUndefined();
		expect(dxa.widthTwips).toBe(7200);
		expect(dxa.widthPercent).toBeUndefined();

		const pct = setTableWidth(dxa, { type: 'pct', value: 0.75 });
		expect(pct.widthTwips).toBeUndefined();
		expect(pct.widthPercent).toBe(0.75);

		const auto = setTableWidth(pct, { type: 'auto' });
		expect(auto.widthTwips).toBeUndefined();
		expect(auto.widthPercent).toBeUndefined();

		expect(() => setTableWidth(base, { type: 'dxa', value: -10 })).toThrow(RangeError);
		expect(() => setTableWidth(base, { type: 'pct', value: 0 })).toThrow(RangeError);
		expect(() => setTableWidth(base, { type: 'pct', value: Number.NaN })).toThrow(RangeError);
	});

	it('updates and clears alignment, indent, style, borders and look flags', () => {
		const base = makeTable();
		const aligned = setTableAlignment(base, 'center');
		expect(aligned.alignment).toBe('center');
		expect(aligned.justification).toBe('center');
		expect(setTableAlignment(aligned, undefined).alignment).toBeUndefined();
		expect(setTableAlignment(aligned, undefined).justification).toBeUndefined();

		const indented = setTableIndent(base, -180);
		expect(indented.indentTwips).toBe(-180);
		expect(setTableIndent(indented, undefined).indentTwips).toBeUndefined();
		expect(() => setTableIndent(base, Number.POSITIVE_INFINITY)).toThrow(RangeError);

		const styled = setTableStyle(base, '  GridTable4-Accent1  ');
		expect(styled.style).toBe('GridTable4-Accent1');
		expect(setTableStyle(styled, '   ').style).toBeUndefined();

		const bordered = setTableBorders(base, {
			top: { style: 'single', sizeEighthPoints: eighthPoints(8), color: '#112233' },
			insideH: { style: 'dashed', sizeEighthPoints: eighthPoints(4), color: '#445566' },
		});
		expect(bordered.borders?.top?.color).toBe('#112233');
		expect(bordered.borders?.insideH?.style).toBe('dashed');
		const narrowed = setTableBorders(bordered, { top: undefined });
		expect(narrowed.borders?.top).toBeUndefined();
		expect(narrowed.borders?.insideH?.style).toBe('dashed');
		expect(setTableBorders(narrowed, { insideH: undefined }).borders).toBeUndefined();
		expect(setTableBorders(bordered, undefined).borders).toBeUndefined();

		const looked = setTableLook(base, { firstRow: true, noHBand: true, lastColumn: false });
		expect(looked.look).toEqual({ firstRow: true, noHBand: true, lastColumn: false });
		const toggled = setTableLook(looked, { noHBand: false, lastColumn: undefined });
		expect(toggled.look).toEqual({ firstRow: true, noHBand: false });
		expect(setTableLook(toggled, undefined).look).toBeUndefined();
	});

	it('updates grid widths and keeps merged cell span widths consistent', () => {
		const merged: Table = {
			type: 'table',
			id: 'tbl-merged',
			grid: [twips(1000), twips(1000), twips(1000)],
			rows: [
				[
					{
						gridSpan: 2,
						widthTwips: twips(2000),
						paragraphs: [{ type: 'paragraph', id: 'm0', runs: [{ text: 'Header' }] }],
					},
					{
						widthTwips: twips(1000),
						paragraphs: [{ type: 'paragraph', id: 'm1', runs: [{ text: 'Right' }] }],
					},
				],
				[
					{
						widthPercent: 0.33,
						paragraphs: [{ type: 'paragraph', id: 'm2', runs: [{ text: 'A' }] }],
					},
					{
						widthTwips: twips(1000),
						paragraphs: [{ type: 'paragraph', id: 'm3', runs: [{ text: 'B' }] }],
					},
					{
						widthTwips: twips(1000),
						paragraphs: [{ type: 'paragraph', id: 'm4', runs: [{ text: 'C' }] }],
					},
				],
			],
		};
		const resized = setTableGridWidths(merged, [1500, 2500, 1200]);
		expect(resized.grid).toEqual([1500, 2500, 1200]);
		expect(resized.rows[0]![0]!.widthTwips).toBe(4000);
		expect(resized.rows[0]![1]!.widthTwips).toBe(1200);
		expect(resized.rows[1]![0]!.widthTwips).toBe(1500);
		expect(resized.rows[1]![0]!.widthPercent).toBeUndefined();
		expect(resized.rows[1]![1]!.widthTwips).toBe(2500);
		expect(resized.rows[1]![2]!.widthTwips).toBe(1200);
		// Original untouched
		expect(merged.grid).toEqual([1000, 1000, 1000]);
		expect(merged.rows[0]![0]!.widthTwips).toBe(2000);

		expect(() => setTableGridWidths(merged, [1500, 2500])).toThrow(RangeError);
		expect(() => setTableGridWidths(merged, [1500, 0, 1200])).toThrow(RangeError);
	});

	it('applies a combined batch with updateTableProperties', () => {
		const updated = updateTableProperties(makeTable(), {
			width: { type: 'dxa', value: 6400 },
			alignment: 'right',
			indentTwips: 240,
			style: 'TableGrid',
			borders: {
				top: { style: 'double', sizeEighthPoints: eighthPoints(12), color: '#AA0000' },
			},
			look: { firstRow: true, lastRow: true, noVBand: true },
			gridWidths: [2400, 4000],
		});
		expect(updated.widthTwips).toBe(6400);
		expect(updated.alignment).toBe('right');
		expect(updated.indentTwips).toBe(240);
		expect(updated.style).toBe('TableGrid');
		expect(updated.borders?.top?.style).toBe('double');
		expect(updated.look).toEqual({ firstRow: true, lastRow: true, noVBand: true });
		expect(updated.grid).toEqual([2400, 4000]);
	});
});

describe('table-properties-edit saveDocx / loadDocx round-trip', () => {
	it('persists all table-level properties on fresh write (saveDocx -> loadDocx)', async () => {
		const doc = createDocument();
		const edited = updateTableProperties(makeTable(), {
			width: { type: 'dxa', value: 8400 },
			alignment: 'center',
			indentTwips: 360,
			style: 'CustomGrid',
			borders: {
				top: { style: 'single', sizeEighthPoints: eighthPoints(8), color: '#112233' },
				bottom: { style: 'double', sizeEighthPoints: eighthPoints(12), color: '#445566' },
				left: { style: 'dashed', sizeEighthPoints: eighthPoints(4), color: '#778899' },
				right: { style: 'dotted', sizeEighthPoints: eighthPoints(6), color: '#AABBCC' },
				insideH: { style: 'single', sizeEighthPoints: eighthPoints(4), color: '#222222' },
				insideV: { style: 'single', sizeEighthPoints: eighthPoints(4), color: '#333333' },
			},
			look: {
				firstRow: true,
				lastRow: true,
				firstColumn: false,
				lastColumn: true,
				noHBand: true,
				noVBand: false,
			},
			gridWidths: [3400, 5000],
		});
		doc.blocks = [edited];
		const bytes = await saveDocx(doc);
		const reopened = expectTable((await loadDocx(bytes)).model.blocks[0]);
		expect(reopened.widthTwips).toBe(8400);
		expect(reopened.widthPercent).toBeUndefined();
		expect(reopened.alignment).toBe('center');
		expect(reopened.indentTwips).toBe(360);
		expect(reopened.style).toBe('CustomGrid');
		expect(reopened.borders).toEqual({
			top: { style: 'single', sizeEighthPoints: 8, color: '#112233' },
			bottom: { style: 'double', sizeEighthPoints: 12, color: '#445566' },
			left: { style: 'dashed', sizeEighthPoints: 4, color: '#778899' },
			right: { style: 'dotted', sizeEighthPoints: 6, color: '#AABBCC' },
			insideH: { style: 'single', sizeEighthPoints: 4, color: '#222222' },
			insideV: { style: 'single', sizeEighthPoints: 4, color: '#333333' },
		});
		expect(reopened.look).toEqual({
			firstRow: true,
			lastRow: true,
			firstColumn: false,
			lastColumn: true,
			noHBand: true,
			noVBand: false,
		});
		expect(reopened.grid).toEqual([3400, 5000]);
		expect(reopened.rows[0]![0]!.widthTwips).toBe(3400);
		expect(reopened.rows[0]![1]!.widthTwips).toBe(5000);
	});

	it('persists percentage and auto table widths on fresh write', async () => {
		const docPct = createDocument();
		docPct.blocks = [setTableWidth(makeTable(), { type: 'pct', value: 0.8 })];
		const pctTable = expectTable((await loadDocx(await saveDocx(docPct))).model.blocks[0]);
		expect(pctTable.widthPercent).toBeCloseTo(0.8);
		expect(pctTable.widthTwips).toBeUndefined();

		const docAuto = createDocument();
		docAuto.blocks = [setTableWidth(makeTable(), { type: 'auto' })];
		const autoBytes = await saveDocx(docAuto);
		expect(await readDocumentXml(autoBytes)).toContain('<w:tblW w:w="0" w:type="auto"/>');
		const autoTable = expectTable((await loadDocx(autoBytes)).model.blocks[0]);
		expect(autoTable.widthTwips).toBeUndefined();
		expect(autoTable.widthPercent).toBeUndefined();
	});

	it('patches existing table-level properties while preserving unrelated XML', async () => {
		const source = await zipXml(`<w:tbl>
<w:tblPr>
<w:tblStyle w:val="OldStyle"/>
<w:tblW w:w="4000" w:type="dxa"/>
<w:jc w:val="left"/>
<w:tblInd w:w="120" w:type="dxa"/>
<w:tblBorders>
<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>
<w:start w:val="single" w:sz="4" w:space="0" w:color="000000"/>
<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>
</w:tblBorders>
<w:tblLayout w:type="fixed"/>
<w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>
<w:tblDescription w:val="Preserve me"/>
</w:tblPr>
<w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>
<w:tr><w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/><w:noWrap/></w:tcPr><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc></w:tr>
</w:tbl>`);
		const loaded = await loadDocx(source);
		const model = structuredClone(loaded.model);
		const table = expectTable(model.blocks[0]);
		model.blocks[0] = updateTableProperties(table, {
			width: { type: 'pct', value: 0.6 },
			alignment: 'right',
			indentTwips: signedTwips(240),
			style: 'NewStyle',
			borders: {
				left: { style: 'thick', sizeEighthPoints: eighthPoints(16), color: '#FF0000' },
				insideV: { style: 'dotted', sizeEighthPoints: eighthPoints(4), color: '#00FF00' },
			},
			look: { lastRow: true, noHBand: true },
			gridWidths: [1500, 3500],
		});
		const saved = await loaded.save(model);
		const xml = await readDocumentXml(saved);
		// Preserved untouched source XML
		expect(xml).toContain('<w:tblLayout w:type="fixed"/>');
		expect(xml).toContain('<w:tblDescription w:val="Preserve me"/>');
		expect(xml).toContain('<w:noWrap/>');
		// Replaced w:start alias cleanly with w:left
		expect(xml).not.toContain('<w:start');
		expect(xml).toContain('<w:tblW w:w="3000" w:type="pct"/>');
		expect(xml).toContain('<w:tblStyle w:val="NewStyle"/>');
		expect(xml).toContain('<w:jc w:val="right"/>');
		expect(xml).toContain('<w:tblInd w:w="240" w:type="dxa"/>');

		const reopened = expectTable((await loadDocx(saved)).model.blocks[0]);
		expect(reopened.widthTwips).toBeUndefined();
		expect(reopened.widthPercent).toBeCloseTo(0.6);
		expect(reopened.alignment).toBe('right');
		expect(reopened.indentTwips).toBe(240);
		expect(reopened.style).toBe('NewStyle');
		expect(reopened.borders?.top?.style).toBe('single');
		expect(reopened.borders?.bottom?.style).toBe('single');
		expect(reopened.borders?.left).toMatchObject({
			style: 'thick',
			sizeEighthPoints: 16,
			color: '#FF0000',
		});
		expect(reopened.borders?.insideV).toMatchObject({
			style: 'dotted',
			sizeEighthPoints: 4,
			color: '#00FF00',
		});
		expect(reopened.look).toEqual({
			firstRow: true,
			lastRow: true,
			firstColumn: true,
			lastColumn: false,
			noHBand: true,
			noVBand: true,
		});
		expect(reopened.grid).toEqual([1500, 3500]);
		expect(reopened.rows[0]![0]!.widthTwips).toBe(1500);
		expect(reopened.rows[0]![1]!.widthTwips).toBe(3500);
	});

	it('clears table-level properties through source-preserving patch and reload', async () => {
		const source = await zipXml(`<w:tbl>
<w:tblPr>
<w:tblStyle w:val="Grid"/>
<w:tblW w:w="5000" w:type="dxa"/>
<w:jc w:val="center"/>
<w:tblInd w:w="180" w:type="dxa"/>
<w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/></w:tblBorders>
<w:tblLook w:val="04A0" w:firstRow="1"/>
</w:tblPr>
<w:tblGrid><w:gridCol w:w="2500"/><w:gridCol w:w="2500"/></w:tblGrid>
<w:tr><w:tc><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc></w:tr>
</w:tbl>`);
		const loaded = await loadDocx(source);
		const model = structuredClone(loaded.model);
		model.blocks[0] = updateTableProperties(expectTable(model.blocks[0]), {
			width: { type: 'auto' },
			alignment: undefined,
			indentTwips: undefined,
			borders: undefined,
			style: undefined,
			look: undefined,
		});
		const saved = await loaded.save(model);
		const xml = await readDocumentXml(saved);
		expect(xml).toContain('<w:tblW w:w="0" w:type="auto"/>');
		expect(xml).not.toMatch(/<w:(tblStyle|jc|tblInd|tblBorders|tblLook)\b/);

		const reopened = expectTable((await loadDocx(saved)).model.blocks[0]);
		expect(reopened.widthTwips).toBeUndefined();
		expect(reopened.widthPercent).toBeUndefined();
		expect(reopened.alignment).toBeUndefined();
		expect(reopened.indentTwips).toBeUndefined();
		expect(reopened.borders).toBeUndefined();
		expect(reopened.style).toBeUndefined();
		expect(reopened.look).toBeUndefined();
	});

	it('patches grid widths and table-level properties on a merged table', async () => {
		const source = await zipXml(`<w:tbl>
<w:tblPr><w:tblW w:w="6000" w:type="dxa"/><w:jc w:val="left"/><w:tblDescription w:val="Merged"/></w:tblPr>
<w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>
<w:tr>
<w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/><w:gridSpan w:val="2"/></w:tcPr><w:p><w:r><w:t>Span</w:t></w:r></w:p></w:tc>
<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>End</w:t></w:r></w:p></w:tc>
</w:tr>
<w:tr>
<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>1</w:t></w:r></w:p></w:tc>
<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>2</w:t></w:r></w:p></w:tc>
<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>3</w:t></w:r></w:p></w:tc>
</w:tr>
</w:tbl>`);
		const loaded = await loadDocx(source);
		const model = structuredClone(loaded.model);
		model.blocks[0] = updateTableProperties(expectTable(model.blocks[0]), {
			width: { type: 'dxa', value: 7500 },
			alignment: 'center',
			style: 'MergedStyle',
			gridWidths: [1500, 3000, 3000],
		});
		const saved = await loaded.save(model);
		const xml = await readDocumentXml(saved);
		expect(xml).toContain('<w:tblDescription w:val="Merged"/>');

		const reopened = expectTable((await loadDocx(saved)).model.blocks[0]);
		expect(reopened.widthTwips).toBe(7500);
		expect(reopened.alignment).toBe('center');
		expect(reopened.style).toBe('MergedStyle');
		expect(reopened.grid).toEqual([1500, 3000, 3000]);
		expect(reopened.rows[0]![0]!.gridSpan).toBe(2);
		expect(reopened.rows[0]![0]!.widthTwips).toBe(4500);
		expect(reopened.rows[0]![1]!.widthTwips).toBe(3000);
		expect(reopened.rows[1]!.map((cell) => cell.widthTwips)).toEqual([1500, 3000, 3000]);
	});
});

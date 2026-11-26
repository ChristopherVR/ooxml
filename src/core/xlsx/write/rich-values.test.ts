import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { encodePng } from '../../drawingml/raster-encode';
import { getCell } from '../cells';
import { createCalcEngine } from '../formula/engine';
import { translateFormula } from '../formula/transform';
import { createEditSession } from '../edit/session';
import { cellView } from '../layout/cell-view';
import { loadXlsx } from '../read/load';
import { RICH_METADATA_EXT, RICH_NS, RICH_PARTS, RICH_TYPES_NS } from '../rich-values';
import { createWorkbook } from '../workbook';
import { miniPackage, R, ws, X } from './mini-package-fixtures';
import { saveXlsx } from './save';

const PNG = encodePng(2, 1, Uint8Array.of(200, 50, 50, 50, 50, 200));
const A1 = { row: 0, col: 0 };

describe('SpreadsheetML rich values', () => {
	it('writes pictures, preserves original parts and metadata bytes, and edits alt text independently', async () => {
		const book = createWorkbook();
		const session = createEditSession(book);
		expect(session.addCellPicture(0, A1, PNG, 'Red & blue')).toBe(true);
		expect(session.addCellPicture(0, { row: 1, col: 0 }, PNG, 'Second')).toBe(true);
		session.setCellInput(0, 0, 1, '=SEQUENCE(2)');
		session.setCellInput(0, 0, 2, '=IMAGE("https://example.test/image.png")');
		const first = await loadXlsx(await saveXlsx(book));
		expect(cellView(first, 0, 0, 0).picture).toMatchObject({
			altText: 'Red & blue',
			pixelWidth: 2,
			pixelHeight: 1,
			sizing: 0,
		});
		expect(getCell(first.sheets[0]!, 0, 1)?.dynamicArray).toBe(true);
		expect(getCell(first.sheets[0]!, 0, 2)).toMatchObject({ value: { error: '#BLOCKED!' } });
		expect(new TextDecoder().decode(first.source?.parts.get('xl/worksheets/sheet1.xml'))).toMatch(
			/<c r="C1" t="e" vm="\d+"><f>_xlfn\.IMAGE.*<v>#VALUE!<\/v><\/c>/,
		);
		const again = await loadXlsx(await saveXlsx(first));
		const parts = [...first.source!.parts.keys()].filter(
			(name) => name.includes('/richData/') || name === 'xl/metadata.xml',
		);
		expect(parts).toHaveLength(6);
		for (const name of parts)
			expect(again.source!.parts.get(name), name).toEqual(first.source!.parts.get(name));

		const edit = createEditSession(again);
		expect(edit.setCellPictureAltText(0, A1, 'New <alt>')).toBe(true);
		edit.undo();
		expect(cellView(again, 0, 0, 0).picture?.altText).toBe('Red & blue');
		edit.redo();
		edit.setCellValue(0, 1, 0, 42);
		const saved = await loadXlsx(await saveXlsx(again));
		expect(cellView(saved, 0, 0, 0).picture?.altText).toBe('New <alt>');
		expect(getCell(saved.sheets[0]!, 1, 0)).toEqual({ value: 42 });
		expect(getCell(saved.sheets[0]!, 0, 1)?.dynamicArray).toBe(true);
	});

	it('stops showing a metadata picture when its formula recalculates to a scalar', async () => {
		const session = createEditSession(createWorkbook());
		session.addCellPicture(0, A1, PNG, 'Cached picture');
		const book = await loadXlsx(await saveXlsx(session.workbook));
		const cell = getCell(book.sheets[0]!, 0, 0)!;
		cell.formula = '1+1';
		expect(cell.valueMetadata).toBeDefined();
		expect(cellView(book, 0, 0, 0).picture?.altText).toBe('Cached picture');
		createCalcEngine(book).recalculateAll();
		expect(cell.value).toBe(2);
		expect(cellView(book, 0, 0, 0).picture).toBeUndefined();
	});

	it('converts a picture in both directions and carries its alt text and media through save', async () => {
		const session = createEditSession(createWorkbook());
		session.addCellPicture(0, A1, PNG, 'Caption');
		expect(session.placePictureOverCells(0, A1)).toBe(0);
		const saved = await loadXlsx(await saveXlsx(session.workbook));
		expect(saved.sheets[0]?.drawings[0]).toMatchObject({ kind: 'image', description: 'Caption' });
		const edit = createEditSession(saved);
		expect(edit.placePictureInCell(0, 0)).toEqual(A1);
		expect(saved.sheets[0]?.drawings).toEqual([]);
		edit.undo();
		expect(saved.sheets[0]?.drawings).toHaveLength(1);
		edit.redo();
		const reloaded = await loadXlsx(await saveXlsx(saved));
		expect(cellView(reloaded, 0, 0, 0).picture?.altText).toBe('Caption');
		expect(reloaded.sheets[0]?.drawings).toHaveLength(0);
	});

	it('shows cached entity fields and preserves unknown data, flags and namespace prefixes', async () => {
		const parts = {
			'xl/metadata.xml': `\ufeff<metadata xmlns="${X}" xmlns:rvalue="${RICH_NS}"><metadataTypes count="1"><metadataType name="XLRICHVALUE" minSupportedVersion="120000"/></metadataTypes><futureMetadata name="XLRICHVALUE" count="1"><bk><extLst><ext uri="${RICH_METADATA_EXT}"><rvalue:rvb i="1"/></ext></extLst></bk></futureMetadata><valueMetadata count="1"><bk><rc t="1" v="0"/></bk></valueMetadata></metadata>`,
			'xl/richData/rdrichvalue.xml': `<r:rvData xmlns:r="${RICH_NS}" count="999999999"><r:rv s="0"><r:v>Seattle</r:v><r:v>WA</r:v><r:v>755078</r:v><r:v>secret</r:v></r:rv><r:rv s="1"><r:v>0</r:v></r:rv><r:rv s="4294967295"><r:v>unknown</r:v></r:rv></r:rvData>`,
			'xl/richData/rdrichvaluestructure.xml': `<rvStructures xmlns="${RICH_NS}"><s t="_linkedentitycore"><k n="_DisplayString" t="s"/><k n="State" t="s"/><k n="人口" t="i"/><k n="Internal" t="s"/></s><s t="_linkedentity"><k n="%cvi" t="r"/></s></rvStructures>`,
			'xl/richData/rdRichValueTypes.xml': `<rvTypesInfo xmlns="${RICH_TYPES_NS}"><global><keyFlags><key name="Internal"><flag name="ShowInCardView" value="0"/></key></keyFlags></global></rvTypesInfo>`,
		};
		const links = Object.values(RICH_PARTS)
			.filter((part) => part !== RICH_PARTS.relationships)
			.map(
				(part, i) =>
					`<Relationship Id="rich${i}" Type="${part.rel}" Target="richData/${part.file}"/>`,
			)
			.join('');
		const book = await loadXlsx(
			await miniPackage({
				sheets: [
					{ name: 'Data', xml: ws('<row r="1"><c r="A1" vm="1" t="e"><v>#VALUE!</v></c></row>') },
				],
				workbookRels: `<Relationship Id="meta" Type="${R}/sheetMetadata" Target="metadata.xml"/>${links}`,
				parts,
			}),
		);
		expect(cellView(book, 0, 0, 0)).toMatchObject({
			text: 'Seattle',
			isError: false,
			linkedData: {
				kind: 'generic',
				fields: [
					{ name: 'State', type: 's', value: 'WA' },
					{ name: '人口', type: 'i', value: '755078' },
				],
			},
		});
		const edit = createEditSession(book);
		edit.setCellInput(0, 0, 1, '=A1.[人口]+1');
		edit.setCellInput(0, 0, 2, `=FIELDVALUE('Data'!$A$1,"state")`);
		edit.setCellInput(0, 0, 3, '=_FV(A1,"Missing")');
		edit.setCellInput(0, 0, 4, '=FIELDVALUE(A1,"Internal")');
		expect(translateFormula('A1.State', 1, 0)).toBe('A2.State');
		const sheet = book.sheets[0]!;
		expect(getCell(sheet, 0, 1)?.value).toBe(755079);
		expect(getCell(sheet, 0, 2)?.value).toBe('WA');
		expect(getCell(sheet, 0, 3)?.value).toEqual({ error: '#FIELD!' });
		expect(getCell(sheet, 0, 4)?.value).toEqual({ error: '#FIELD!' });
		const bytes = await saveXlsx(book);
		const reopened = await loadXlsx(bytes);
		const reopenedSheet = reopened.sheets[0]!;
		expect(getCell(reopenedSheet, 0, 1)?.formula).toBe('_FV(A1,"人口")+1');
		const reopenedEdit = createEditSession(reopened);
		reopenedEdit.calculateNow({ full: true });
		expect(getCell(reopenedSheet, 0, 1)?.value).toBe(755079);
		reopenedEdit.setCellValue(0, 0, 0, 'ordinary text');
		expect(getCell(reopenedSheet, 0, 1)?.value).toEqual({ error: '#FIELD!' });
		expect(getCell(reopenedSheet, 0, 2)?.value).toEqual({ error: '#FIELD!' });
		const saved = await JSZip.loadAsync(bytes);
		for (const [name, xml] of Object.entries(parts))
			expect(await saved.file(name)?.async('uint8array'), name).toEqual(
				new TextEncoder().encode(xml),
			);
	});
});

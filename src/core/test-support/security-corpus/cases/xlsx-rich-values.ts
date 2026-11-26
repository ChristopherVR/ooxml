import { encodePng } from '../../../drawingml/raster-encode';
import { createEditSession } from '../../../xlsx/edit/session';
import { createCalcEngine } from '../../../xlsx/formula/engine';
import { book, get } from '../../../xlsx/formula/test-helpers';
import { cellView } from '../../../xlsx/layout/cell-view';
import { loadXlsx } from '../../../xlsx/read/load';
import { RICH_NS, RICH_PARTS } from '../../../xlsx/rich-values';
import { createWorkbook } from '../../../xlsx/workbook';
import { miniPackage, R, X, ws } from '../../../xlsx/write/mini-package-fixtures';
import type { SecurityCase } from '../case';

export const cases: readonly SecurityCase[] = [
	{
		name: 'An oversized decoded in-cell picture is refused before reaching an image decoder',
		limits: { ms: 2500, heapMb: 64 },
		run() {
			const png = encodePng(1, 1, Uint8Array.of(0, 0, 0));
			new DataView(png.buffer).setUint32(16, 0xffffffff);
			return createEditSession(createWorkbook()).addCellPicture(0, { row: 0, col: 0 }, png);
		},
		expected: false,
	},
	{
		name: 'An external rich-data picture relationship stays inert',
		limits: { ms: 2500, heapMb: 64 },
		async run() {
			const { data, structures, relationships } = RICH_PARTS;
			const rels = [data, structures, relationships]
				.map(
					(part, i) =>
						`<Relationship Id="rd${i}" Type="${part.rel}" Target="richData/${part.file}"/>`,
				)
				.join('');
			const workbook = await loadXlsx(
				await miniPackage({
					sheets: [
						{ name: 'Data', xml: ws('<row r="1"><c r="A1" vm="1" t="e"><v>#VALUE!</v></c></row>') },
					],
					workbookRels: `<Relationship Id="meta" Type="${R}/sheetMetadata" Target="metadata.xml"/>${rels}`,
					parts: {
						'xl/metadata.xml': `<metadata xmlns="${X}" xmlns:rv="${RICH_NS}"><metadataTypes><metadataType name="XLRICHVALUE"/></metadataTypes><futureMetadata name="XLRICHVALUE"><bk><extLst><ext><rv:rvb i="0"/></ext></extLst></bk></futureMetadata><valueMetadata><bk><rc t="1" v="0"/></bk></valueMetadata></metadata>`,
						[`xl/richData/${data.file}`]: `<rvData xmlns="${RICH_NS}" count="4294967295"><rv s="0"><v>0</v></rv></rvData>`,
						[`xl/richData/${structures.file}`]: `<rvStructures xmlns="${RICH_NS}"><s t="_localImage"><k n="_rvRel:LocalImageIdentifier" t="i"/></s></rvStructures>`,
						[`xl/richData/${relationships.file}`]: `<richValueRels xmlns="${relationships.ns}" xmlns:r="${R}"><rel r:id="remote"/></richValueRels>`,
						[`xl/richData/_rels/${relationships.file}.rels`]: `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="remote" Type="${R}/image" Target="https://example.test/picture.png" TargetMode="External"/></Relationships>`,
					},
				}),
			);
			const cell = cellView(workbook, 0, 0, 0);
			return [workbook.richValues?.records.length, cell.picture === undefined, cell.text];
		},
		expected: [1, true, '#VALUE!'],
	},
	{
		name: 'A workbook formula cannot initiate an image request by default',
		limits: { ms: 2500, heapMb: 64 },
		run() {
			const workbook = book({ A1: '=IMAGE("https://example.test/picture.png")' });
			createCalcEngine(workbook).recalculateAll();
			return get(workbook, 'A1');
		},
		expected: { error: '#BLOCKED!' },
	},
];

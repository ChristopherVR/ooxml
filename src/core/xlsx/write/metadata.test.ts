import { describe, expect, it } from 'vitest';
import { RELATIONSHIP_TYPES } from '../../opc/index.js';
import type { Cell } from '../model.js';
import { loadXlsx } from '../read/index.js';
import { parseDynamicArrayMetadata } from '../read/metadata.js';
import { createWorksheet } from '../workbook.js';
import { MetadataPlan, writeMetadataPart } from './metadata.js';
import { PackageWriter, RelationshipSet } from './package-writer.js';
import { miniPackage, ws, X } from './mini-package-fixtures.js';
import { SharedStringTable } from './shared-strings.js';
import { sheetDataXml } from './sheet-data.js';

const RICH = 'http://schemas.microsoft.com/office/spreadsheetml/2017/richdata';
// The review's t12 metadata: one XLRICHVALUE value-metadata block, no cell metadata.
const RICH_METADATA = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<metadata xmlns="${X}" xmlns:xlrd="${RICH}"><metadataTypes count="1"><metadataType name="XLRICHVALUE" minSupportedVersion="120000" copy="1" pasteAll="1" pasteValues="1" merge="1" splitFirst="1" rowColShift="1" clearFormats="1" clearComments="1" assign="1" coerce="1"/></metadataTypes><futureMetadata name="XLRICHVALUE" count="1"><bk><extLst><ext uri="{3e2802c4-a4d2-4d8b-9148-e3be6c30e623}"><xlrd:rvb i="0"/></ext></extLst></bk></futureMetadata><valueMetadata count="1"><bk><rc t="1" v="0"/></bk></valueMetadata></metadata>`;

describe('MetadataPlan', () => {
	it('keeps a source part unchanged when no dynamic array is written', () => {
		const plan = new MetadataPlan(RICH_METADATA);
		expect(plan.keepsSource).toBe(true);
		expect(plan.xml(false)).toBe(RICH_METADATA);
	});

	it('merges the XLDAPR block into the source, keeping value metadata indices', () => {
		const plan = new MetadataPlan(RICH_METADATA);
		const xml = plan.xml(true) ?? '';
		expect(plan.dynamicCm).toBe(1);
		expect(parseDynamicArrayMetadata(xml)).toEqual(new Set([1]));
		expect(xml).toContain('<valueMetadata count="1"><bk><rc t="1" v="0"/></bk></valueMetadata>');
		expect(xml).toContain('<metadataTypes count="2">');
		expect(xml).toMatch(
			/<cellMetadata count="1"><bk><rc t="2" v="0"\/><\/bk><\/cellMetadata><valueMetadata/,
		);
		expect(xml).toContain('<xlrd:rvb i="0"/>');
	});

	it('appends after existing cell metadata blocks and reuses an existing XLDAPR block', () => {
		const withCells = RICH_METADATA.replace(
			'<valueMetadata',
			'<cellMetadata count="2"><bk><rc t="1" v="0"/></bk><bk><rc t="1" v="0"/></bk></cellMetadata><valueMetadata',
		);
		const plan = new MetadataPlan(withCells);
		expect(plan.dynamicCm).toBe(3);
		const merged = plan.xml(true) ?? '';
		expect(parseDynamicArrayMetadata(merged)).toEqual(new Set([3]));
		const again = new MetadataPlan(merged);
		expect(again.dynamicCm).toBe(3);
		expect(again.xml(true)).toBe(merged);
	});

	it('writes the default part only for new dynamic arrays when there is no source', () => {
		expect(new MetadataPlan(undefined).xml(false)).toBeUndefined();
		expect(parseDynamicArrayMetadata(new MetadataPlan(undefined).xml(true))).toEqual(new Set([1]));
	});

	it('writes the part at the source name with a workbook relationship', () => {
		const writer = new PackageWriter(undefined);
		const rels = new RelationshipSet();
		const plan = new MetadataPlan(RICH_METADATA);
		writeMetadataPart(writer, rels, 'xl/workbook.xml', plan, false, 'xl/metadata.xml', new Set());
		expect(writer.has('xl/metadata.xml')).toBe(true);
		expect([...rels.entries.values()]).toEqual([
			{ type: RELATIONSHIP_TYPES.sheetMetadata, target: 'metadata.xml', mode: 'Internal' },
		]);
	});
});

describe('cell vm', () => {
	it('reads vm and writes it back only while the metadata part is kept', async () => {
		const wb = await loadXlsx(
			await miniPackage({
				sheets: [
					{ name: 'S', xml: ws('<row r="1"><c r="A1" t="e" vm="1"><v>#VALUE!</v></c></row>') },
				],
			}),
		);
		const cell = wb.sheets[0]?.rows.get(0)?.get(0);
		expect(cell?.valueMetadata).toBe(1);
		const sheet = createWorksheet('S', 1);
		sheet.rows.set(
			0,
			new Map<number, Cell>([
				[0, cell!],
				[1, { ...cell!, value: 5 }],
			]),
		);
		const base = { strings: new SharedStringTable(), styleCount: 1, headerText: new Map() };
		expect(sheetDataXml(base, sheet)).not.toContain('vm=');
		const kept = sheetDataXml({ ...base, metadata: new MetadataPlan(RICH_METADATA) }, sheet);
		expect(kept).toContain('<c r="A1" t="e" vm="1"><v>#VALUE!</v></c>');
		// A typed value no longer points at the rich value.
		expect(kept).toContain('<c r="B1"><v>5</v></c>');
	});
});

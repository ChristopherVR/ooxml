import type { PptxTableData } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import { translationsEn } from '../../i18n/translations-en';
import { inspectorFieldPatch, resolveInspectorSections } from './inspector-field-model';
import { TABLE_OPTIONS_SECTIONS } from './table-option-fields';

const table = (extra: Partial<PptxTableData> = {}): PptxTableData =>
	({
		columnWidths: [100, 100],
		rows: [{ height: 20, cells: [{ text: 'a', style: { marginLeft: 8 } }, { text: 'b' }] }],
		...extra,
	}) as PptxTableData;

const patch = (data: PptxTableData, field: string, raw: string | boolean) =>
	inspectorFieldPatch(TABLE_OPTIONS_SECTIONS, 'tableOptions', field, data, raw);

describe('table option fields', () => {
	it("shows PowerPoint's style options and a band size only for a banded table", () => {
		const [plain] = resolveInspectorSections(TABLE_OPTIONS_SECTIONS, table());
		expect(plain!.fields.map((f) => f.id)).toStrictEqual([
			'firstRowHeader',
			'firstCol',
			'lastRow',
			'lastCol',
			'bandedRows',
			'bandedColumns',
			'rtl',
			'cellMargin',
		]);
		expect(plain!.fields.at(-1)!.value).toBe(6);
		const [banded] = resolveInspectorSections(
			TABLE_OPTIONS_SECTIONS,
			table({ bandedRows: true, bandRowCycle: 2 }),
		);
		expect(banded!.fields.find((f) => f.id === 'bandRowCycle')?.value).toBe(2);
	});

	it('writes table data patches', () => {
		expect(patch(table(), 'lastRow', true)).toStrictEqual({ lastRow: true });
		expect(patch(table({ bandedRows: true }), 'bandRowCycle', '0')).toStrictEqual({
			bandRowCycle: 1,
		});
		const margin = patch(table(), 'cellMargin', '3') as Partial<PptxTableData>;
		expect(margin.rows?.[0]?.cells.map((cell) => cell.style?.marginTop)).toStrictEqual([4, 4]);
	});

	it('labels every field with a dictionary key', () => {
		const keys = TABLE_OPTIONS_SECTIONS[0]!.fields.map((f) => f.labelKey);
		expect(keys.filter((key) => !(key in translationsEn))).toStrictEqual([]);
	});
});

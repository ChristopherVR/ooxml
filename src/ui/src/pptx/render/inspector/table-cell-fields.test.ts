import { describe, expect, it } from 'vitest';

import { translationsEn } from '../../i18n/translations-en';
import { inspectorFieldPatch, resolveInspectorSections } from './inspector-field-model';
import { TABLE_CELL_SECTIONS } from './table-cell-fields';

describe('table cell fields', () => {
	it('shows the text options, then a titled border section per side', () => {
		const views = resolveInspectorSections(TABLE_CELL_SECTIONS, { borderTopWidth: 4 });
		expect(views.map((v) => [v.id, v.labelKey ?? ''])).toStrictEqual([
			['tableCellText', ''],
			['tableCellMargins', 'pptx.table.margins'],
			['tableCellBorderTop', 'pptx.table.borderTop'],
			['tableCellBorderBottom', 'pptx.table.borderBottom'],
			['tableCellBorderLeft', 'pptx.table.borderLeft'],
			['tableCellBorderRight', 'pptx.table.borderRight'],
			['tableCellBorderDiagDown', 'pptx.table.borderDiagDown'],
			['tableCellBorderDiagUp', 'pptx.table.borderDiagUp'],
		]);
		expect(views[2]!.fields.map((f) => [f.id, f.value])).toStrictEqual([
			['color', '#374151'],
			['width', 3],
		]);
	});

	it('writes cell style patches in the model units', () => {
		const patch = (section: string, field: string, raw: string | boolean) =>
			inspectorFieldPatch(TABLE_CELL_SECTIONS, section, field, {}, raw);
		expect(patch('tableCellText', 'bold', true)).toStrictEqual({ bold: true });
		expect(patch('tableCellText', 'vAlign', 'middle')).toStrictEqual({ vAlign: 'middle' });
		expect(patch('tableCellMargins', 'marginTop', '6')).toStrictEqual({ marginTop: 8 });
		expect(patch('tableCellBorderLeft', 'width', '1.5')).toStrictEqual({ borderLeftWidth: 2 });
		expect(patch('tableCellBorderLeft', 'color', '#ff0000')).toStrictEqual({
			borderLeftColor: '#ff0000',
		});
	});

	it('labels every section, field and choice with a dictionary key', () => {
		const keys = TABLE_CELL_SECTIONS.flatMap((s) => [
			...(s.labelKey ? [s.labelKey] : []),
			...s.fields.flatMap((f) => [f.labelKey, ...(f.options ?? []).map((o) => o.labelKey)]),
		]);
		expect(keys.filter((key) => !(key in translationsEn))).toStrictEqual([]);
	});
});

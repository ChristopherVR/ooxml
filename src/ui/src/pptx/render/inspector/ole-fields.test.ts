import type { PptxElement } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import { translationsEn } from '../../i18n/translations-en';
import { inspectorFieldPatch, resolveInspectorSections } from './inspector-field-model';
import { OLE_OBJECT_NAME_SECTIONS, resolveOleSummaryRows } from './ole-fields';

const ole = (extra: Record<string, unknown> = {}): PptxElement =>
	({ type: 'ole', id: 'o1', x: 0, y: 0, width: 1, height: 1, ...extra }) as PptxElement;

describe('ole card', () => {
	it('edits the Object Name as typed and removes a blank one', () => {
		const [view] = resolveInspectorSections(OLE_OBJECT_NAME_SECTIONS, ole({ oleName: 'Budget' }));
		expect(view!.fields[0]!.value).toBe('Budget');
		const write = (raw: string) =>
			inspectorFieldPatch(OLE_OBJECT_NAME_SECTIONS, 'oleObject', 'oleName', ole(), raw);
		expect(write('Q3 ')).toStrictEqual({ oleName: 'Q3 ' });
		expect(write('')).toStrictEqual({ oleName: undefined });
		expect(write('   ')).toStrictEqual({ oleName: undefined });
	});

	it('lists the read-only rows, the file name only when there is one', () => {
		expect(resolveOleSummaryRows(ole({ isLinked: true })).map((r) => r.id)).toStrictEqual([
			'type',
			'linkStatus',
		]);
		const rows = resolveOleSummaryRows(ole({ fileName: 'book.xlsx' }));
		expect(rows.map((r) => [r.id, r.value ?? r.valueKey])).toStrictEqual([
			['type', rows[0]!.value],
			['fileName', 'book.xlsx'],
			['linkStatus', 'pptx.ole.embedded'],
		]);
		expect(resolveOleSummaryRows({ type: 'shape' } as PptxElement)).toStrictEqual([]);
	});

	it('labels every field and row with a dictionary key', () => {
		const keys = [
			...OLE_OBJECT_NAME_SECTIONS[0]!.fields.flatMap((f) => [f.labelKey, f.placeholderKey ?? '']),
			...resolveOleSummaryRows(ole({ fileName: 'x' })).flatMap((r) => [
				r.labelKey,
				...(r.valueKey ? [r.valueKey] : []),
			]),
		];
		expect(keys.filter((key) => !(key in translationsEn))).toStrictEqual([]);
	});
});

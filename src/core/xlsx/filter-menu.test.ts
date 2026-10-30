import { describe, expect, it } from 'vitest';
import { parseRange } from './address';
import { createEditSession } from './edit/session';
import {
	betweenFilter,
	checklistCriteria,
	colorFilterCriteria,
	customFilter,
	decodeCondition,
	encodeCondition,
	filterColumnKind,
	filterMenuModel,
} from './filter-menu';
import { createWorkbook } from './workbook';

function setup() {
	const workbook = createWorkbook();
	const session = createEditSession(workbook, { autoRowHeight: false });
	session.setRangeValues(0, { row: 0, col: 0 }, [
		['Kind', 'Amount'],
		['b', 30],
		['a', 10],
		['b', 20],
		[null, 40],
	]);
	session.setAutoFilter(0, parseRange('A1:B5')!);
	return { workbook, session, sheet: workbook.sheets[0]! };
}

const hidden = (sheet: ReturnType<typeof setup>['sheet']) =>
	[1, 2, 3, 4].filter((r) => sheet.rowInfo.get(r)?.hidden);

describe('filter menu model', () => {
	it('lists distinct values, blanks and the column kind', () => {
		const { workbook, sheet } = setup();
		const filter = sheet.autoFilter!;
		const model = filterMenuModel(workbook, sheet, filter, 0);
		expect(model.kind).toBe('text');
		expect(model.values.map((v) => [v.text, v.checked])).toEqual([
			['a', true],
			['b', true],
		]);
		expect(model.blanks).toEqual({ checked: true });
		expect(filterColumnKind(workbook, sheet, filter, 1)).toBe('number');
		expect(filterMenuModel(workbook, sheet, filter, 1).values.map((v) => v.text)).toEqual([
			'10',
			'20',
			'30',
			'40',
		]);
	});

	it('reflects other columns and the own criteria in the checklist', () => {
		const { workbook, session, sheet } = setup();
		session.setFilterCriteria(0, 1, { top10: { top: true, percent: false, value: 2 } });
		expect(hidden(sheet)).toEqual([2, 3]);
		const model = filterMenuModel(workbook, sheet, sheet.autoFilter!, 0);
		expect(model.values.map((v) => [v.text, v.checked])).toEqual([['b', true]]);
		const own = filterMenuModel(workbook, sheet, sheet.autoFilter!, 1);
		expect(own.values.map((v) => [v.text, v.checked])).toEqual([
			['10', false],
			['20', false],
			['30', true],
			['40', true],
		]);
	});

	it('builds checklist criteria and clears when everything is ticked', () => {
		const { workbook, sheet } = setup();
		const model = filterMenuModel(workbook, sheet, sheet.autoFilter!, 0);
		expect(checklistCriteria(model, new Set(['a', 'b']), true)).toBeUndefined();
		expect(checklistCriteria(model, new Set(['a']), true)).toEqual({ values: ['a'], blank: true });
		expect(checklistCriteria(model, new Set(['b']), false)).toEqual({ values: ['b'] });
	});

	it('encodes and decodes begins, ends and contains with escaped wildcards', () => {
		for (const op of ['beginsWith', 'notEndsWith', 'contains', 'notContains', 'equals'] as const) {
			expect(decodeCondition(encodeCondition(op, 'a*b'))).toEqual({ op, value: 'a*b' });
		}
		expect(encodeCondition('beginsWith', 'x')).toEqual({ operator: 'equal', value: 'x*' });
		expect(encodeCondition('notContains', 'x')).toEqual({ operator: 'notEqual', value: '*x*' });
		expect(betweenFilter('1', '5').custom).toEqual({
			and: true,
			conditions: [
				{ operator: 'greaterThanOrEqual', value: '1' },
				{ operator: 'lessThanOrEqual', value: '5' },
			],
		});
		expect(customFilter([{ op: 'equals', value: 'a' }], true).custom).toEqual({
			conditions: [{ operator: 'equal', value: 'a' }],
		});
	});
});

describe('filtering through the session', () => {
	it('hides rows by custom criteria and clears them', () => {
		const { session, sheet } = setup();
		session.setFilterCriteria(0, 0, customFilter([{ op: 'beginsWith', value: 'b' }], false));
		expect(hidden(sheet)).toEqual([2, 4]);
		session.setFilterCriteria(0, 1, { dynamic: { type: 'aboveAverage' } });
		expect(hidden(sheet)).toEqual([2, 3, 4]);
		session.setFilterCriteria(0, 0, undefined);
		session.setFilterCriteria(0, 1, undefined);
		expect(hidden(sheet)).toEqual([]);
		expect(sheet.autoFilter?.columns).toBeUndefined();
	});

	it('filters by fill colour and reapplies after an edit', () => {
		const { session, sheet } = setup();
		session.applyStyle(0, [parseRange('A2')!], {
			fill: { type: 'pattern', pattern: 'solid', fgColor: { rgb: 'FFFF0000' } },
		});
		session.setFilterCriteria(0, 0, colorFilterCriteria('#FF0000', true));
		expect(hidden(sheet)).toEqual([2, 3, 4]);
		session.setCellValue(0, 2, 0, 'c');
		expect(hidden(sheet)).toEqual([2, 3, 4]);
		session.applyStyle(0, [parseRange('A3')!], {
			fill: { type: 'pattern', pattern: 'solid', fgColor: { rgb: 'FFFF0000' } },
		});
		session.reapplyFilter(0);
		expect(hidden(sheet)).toEqual([3, 4]);
	});
});

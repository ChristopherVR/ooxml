import { describe, expect, it } from 'vitest';
import { formatRange, parseRange, type CellRange } from '../address';
import type { ConditionalFormat, Table, Workbook } from '../model';
import { createWorkbook } from '../workbook';
import { compareWorkbooks } from './compare';

const R = (ref: string): CellRange => {
	const range = parseRange(ref);
	if (!range) throw new Error(ref);
	return range;
};

/** A workbook whose sheets have the given names and `sheetId`s. */
function book(sheets: [name: string, id: number][]): Workbook {
	const workbook = createWorkbook({ sheets: sheets.map(([name]) => name) });
	workbook.sheets.forEach((sheet, i) => {
		sheet.sheetId = sheets[i]![1];
	});
	return workbook;
}

const table = (name: string, ref: string, columns = ['A', 'B']): Table => ({
	id: 1,
	name,
	displayName: name,
	range: R(ref),
	headerRow: true,
	totalsRow: false,
	columns: columns.map((c) => ({ name: c })),
	styleName: 'TableStyleMedium2',
});

const highlight = (ref: string, priority: number): ConditionalFormat => ({
	ranges: [R(ref)],
	rules: [
		{
			type: 'cellIs',
			operator: 'greaterThan',
			formulas: ['10'],
			style: { font: { bold: true } },
			priority,
		},
	],
});

describe('compareWorkbooks structure', () => {
	it('reports sheets removed, added, renamed, moved and hidden', () => {
		const before = book([
			['Data', 1],
			['Old', 2],
			['Calc', 4],
			['Notes', 3],
		]);
		const after = book([
			['Notes', 3],
			['Inputs', 1],
			['Calc', 4],
			['Summary', 5],
		]);
		after.sheets[2]!.state = 'hidden';
		const result = compareWorkbooks(before, after);
		expect(result.workbook).toEqual([
			{ kind: 'sheetRemoved', sheet: 'Old', index: 1 },
			{ kind: 'sheetAdded', sheet: 'Summary', index: 3 },
			{ kind: 'sheetMoved', sheet: 'Notes', from: 3, to: 0 },
			{ kind: 'sheetRenamed', from: 'Data', to: 'Inputs' },
			{ kind: 'sheetState', sheet: 'Calc', from: 'visible', to: 'hidden' },
		]);
		expect(result.sheets.map((s) => [s.beforeName, s.name, s.matchedBy])).toEqual([
			['Notes', 'Notes', 'name'],
			['Data', 'Inputs', 'sheetId'],
			['Calc', 'Calc', 'name'],
		]);
	});

	it('pairs by name first, so renumbered sheet ids do not look like renames', () => {
		const before = book([
			['Alpha', 1],
			['Beta', 2],
		]);
		const renumbered = book([
			['Beta', 1],
			['Alpha', 2],
		]);
		const auto = compareWorkbooks(before, renumbered);
		expect(auto.workbook.map((d) => d.kind)).toEqual(['sheetMoved']);
		expect(auto.sheets.map((s) => s.matchedBy)).toEqual(['name', 'name']);
		const byId = compareWorkbooks(before, renumbered, { sheetMatch: 'sheetId' });
		expect(byId.workbook).toEqual([
			{ kind: 'sheetRenamed', from: 'Alpha', to: 'Beta' },
			{ kind: 'sheetRenamed', from: 'Beta', to: 'Alpha' },
		]);
		const renamed = book([
			['Gamma', 1],
			['Beta', 2],
		]);
		expect(compareWorkbooks(before, renamed, { sheetMatch: 'name' }).workbook).toEqual([
			{ kind: 'sheetRemoved', sheet: 'Alpha', index: 0 },
			{ kind: 'sheetAdded', sheet: 'Gamma', index: 0 },
		]);
	});

	it('reports defined names added, removed and changed, following their sheets', () => {
		const before = book([
			['Data', 1],
			['Calc', 2],
		]);
		before.definedNames = [
			{ name: 'Rate', formula: 'Data!$B$1' },
			{ name: 'Total', formula: 'Calc!$A$1', localSheet: 1 },
			{ name: 'Old', formula: '1' },
			{ name: 'Same', formula: 'Data!$A$1', localSheet: 0 },
		];
		const after = book([
			['Calc', 2],
			['Data', 1],
		]);
		after.definedNames = [
			{ name: 'RATE', formula: 'Data!$B$2' },
			{ name: 'Total', formula: 'Calc!$A$1', localSheet: 0, hidden: true },
			{ name: 'Same', formula: 'Data!$A$1', localSheet: 1 },
			{ name: 'New', formula: '2', localSheet: 1 },
		];
		before.definedNames.push({ name: 'Spelled', formula: '3' });
		after.definedNames.push({ name: 'SPELLED', formula: '3' });
		const names = compareWorkbooks(before, after, { categories: ['name'] }).workbook;
		expect(names.map((d) => ('name' in d ? [d.kind, d.name, d.scope] : [d.kind]))).toEqual([
			['nameAdded', 'New', 'Data'],
			['nameRemoved', 'Old', undefined],
			['nameChanged', 'RATE', undefined],
			['nameChanged', 'SPELLED', undefined],
			['nameChanged', 'Total', 'Calc'],
		]);
	});

	it('reports merges, tables, data validation and conditional formats', () => {
		const before = createWorkbook();
		const after = createWorkbook();
		const [b, a] = [before.sheets[0]!, after.sheets[0]!];
		b.merges = [R('A1:C1'), R('E5:F6')];
		a.merges = [R('E5:F6'), R('A2:B2')];
		b.tables = [table('Sales', 'A10:B20'), table('Costs', 'D10:E20'), table('Gone', 'H1:I2')];
		a.tables = [
			table('Sales', 'A10:B30'),
			table('Costs', 'D10:E20', ['A', 'Total']),
			table('Fresh', 'K1:L2'),
		];
		b.dataValidations = [
			{ ranges: [R('B2'), R('A2')], type: 'list', formula1: '"a,b"' },
			{ ranges: [R('C2')], type: 'whole', operator: 'between', formula1: '1', formula2: '9' },
		];
		a.dataValidations = [
			{ ranges: [R('A2'), R('B2')], type: 'list', formula1: '"a,b"' },
			{ ranges: [R('C2')], type: 'whole', operator: 'between', formula1: '1', formula2: '10' },
		];
		// Renumbered priorities alone are not a change; a new range is.
		b.conditionalFormats = [highlight('A1:A9', 1), highlight('B1:B9', 2)];
		a.conditionalFormats = [highlight('A1:A9', 2), highlight('B1:B20', 1)];
		const structure = compareWorkbooks(before, after).sheets[0]!.structure;
		expect(
			structure.map((d) => {
				if ('range' in d) return `${d.kind} ${formatRange(d.range)}`;
				if (d.kind === 'tableChanged') return `${d.kind} ${d.table} ${d.fields.join()}`;
				if ('table' in d) return `${d.kind} ${d.table}`;
				const item = 'before' in d ? d.before : d.after;
				return `${d.kind} ${item.ranges.map(formatRange).join(' ')}`;
			}),
		).toEqual([
			'mergeRemoved A1:C1',
			'mergeAdded A2:B2',
			'tableChanged Sales range',
			'tableChanged Costs columns',
			'tableRemoved Gone',
			'tableAdded Fresh',
			'dataValidationRemoved C2',
			'dataValidationAdded C2',
			'conditionalFormatRemoved B1:B9',
			'conditionalFormatAdded B1:B20',
		]);
	});

	it('reports a conditional format that moved above or below other rules', () => {
		const before = createWorkbook();
		const after = createWorkbook();
		const formats = ['A1:A9', 'B1:B9', 'C1:C9'];
		before.sheets[0]!.conditionalFormats = formats.map((ref, i) => highlight(ref, i + 1));
		// A new rule on top renumbers every priority without reordering the others.
		after.sheets[0]!.conditionalFormats = [
			highlight('D1:D9', 1),
			...formats.map((ref, i) => highlight(ref, i + 2)),
		];
		const renumbered = compareWorkbooks(before, after).sheets[0]!.structure;
		expect(renumbered.map((d) => d.kind)).toEqual(['conditionalFormatAdded']);
		// C moves above A and B: one move, not three.
		after.sheets[0]!.conditionalFormats = [
			highlight('A1:A9', 2),
			highlight('B1:B9', 3),
			highlight('C1:C9', 1),
		];
		const [moved, ...rest] = compareWorkbooks(before, after).sheets[0]!.structure;
		expect(rest).toEqual([]);
		// The same format twice, and rules stored in another order, are not changes.
		const twice = createWorkbook();
		twice.sheets[0]!.conditionalFormats = [
			highlight('A1:A9', 1),
			highlight('B1:B9', 2),
			highlight('A1:A9', 3),
		];
		expect(compareWorkbooks(twice, structuredClone(twice)).count).toBe(0);
		const pair = (first: number, second: number): ConditionalFormat => ({
			ranges: [R('E1:E9')],
			rules: [
				{ type: 'containsBlanks', style: { font: { bold: true } }, priority: first },
				{ type: 'containsErrors', style: { font: { italic: true } }, priority: second },
			],
		});
		const stored = createWorkbook();
		stored.sheets[0]!.conditionalFormats = [pair(1, 2)];
		const swapped = createWorkbook();
		const [blanks, errors] = pair(1, 2).rules;
		swapped.sheets[0]!.conditionalFormats = [{ ranges: [R('E1:E9')], rules: [errors!, blanks!] }];
		expect(compareWorkbooks(stored, swapped).count).toBe(0);
		// Equal formats pair in priority order, whatever order they are stored in.
		const restored = createWorkbook();
		restored.sheets[0]!.conditionalFormats = [
			highlight('A1:A9', 3),
			highlight('B1:B9', 2),
			highlight('A1:A9', 1),
		];
		expect(compareWorkbooks(twice, restored).count).toBe(0);
		// Colours compare canonically; a differential `bold: false` switches bold off, so it counts.
		const tinted = (style: object): ConditionalFormat => ({
			ranges: [R('F1:F9')],
			rules: [{ type: 'containsErrors', style, priority: 1 }],
		});
		const red = createWorkbook();
		red.sheets[0]!.conditionalFormats = [tinted({ font: { color: { rgb: 'FF0000', tint: 0 } } })];
		const spelled = createWorkbook();
		spelled.sheets[0]!.conditionalFormats = [tinted({ font: { color: { rgb: 'ffff0000' } } })];
		expect(compareWorkbooks(red, spelled).count).toBe(0);
		const off = createWorkbook();
		off.sheets[0]!.conditionalFormats = [
			tinted({ font: { color: { rgb: 'FF0000' }, bold: false } }),
		];
		expect(compareWorkbooks(red, off).count).toBe(2);
		// Rule settings spelled out at their defaults are not changes; turned on, they are.
		const bar = (extra: object): ConditionalFormat => ({
			ranges: [R('G1:G9')],
			rules: [
				{
					type: 'dataBar',
					min: { type: 'min' },
					max: { type: 'max' },
					color: { rgb: '638EC6' },
					priority: 1,
					...extra,
				},
			],
		});
		const plainBar = createWorkbook();
		plainBar.sheets[0]!.conditionalFormats = [bar({})];
		const spelledBar = createWorkbook();
		spelledBar.sheets[0]!.conditionalFormats = [
			bar({ stopIfTrue: false, showValue: true, min: { type: 'min', gte: true } }),
		];
		expect(compareWorkbooks(plainBar, spelledBar).count).toBe(0);
		spelledBar.sheets[0]!.conditionalFormats = [bar({ stopIfTrue: true })];
		expect(compareWorkbooks(plainBar, spelledBar).count).toBe(2);
		// Rules grouped into formats differently, and empty parts of a differential style, are not
		// changes either.
		const rules = (first: object, second: object) => [
			{ type: 'containsBlanks' as const, style: first, priority: 1 },
			{ type: 'containsErrors' as const, style: second, priority: 2 },
		];
		const grouped = createWorkbook();
		grouped.sheets[0]!.conditionalFormats = [
			{ ranges: [R('H1:H9')], rules: rules({ font: { bold: true } }, {}) },
		];
		const split = createWorkbook();
		const [blank, error] = rules({ font: { bold: true } }, { font: {} });
		split.sheets[0]!.conditionalFormats = [
			{ ranges: [R('H1:H9')], rules: [blank!] },
			{ ranges: [R('H1:H9')], rules: [error!] },
		];
		expect(compareWorkbooks(grouped, split).count).toBe(0);
		// Swapping the priorities themselves changes which rule wins.
		swapped.sheets[0]!.conditionalFormats = [pair(2, 1)];
		expect(compareWorkbooks(stored, swapped).sheets[0]!.structure.map((d) => d.kind)).toEqual([
			'conditionalFormatReordered',
		]);
		expect(moved).toMatchObject({
			kind: 'conditionalFormatReordered',
			after: { ranges: [R('C1:C9')] },
		});
	});

	it('keeps the anchor of a rule with relative formulas, not of one without', () => {
		const format = (
			refs: string[],
			rule: ConditionalFormat['rules'][number],
		): ConditionalFormat => ({
			ranges: refs.map(R),
			rules: [rule],
		});
		const style = { font: { bold: true } };
		const expression = { type: 'expression' as const, formula: 'A1=1', style, priority: 1 };
		const blanks = { type: 'containsBlanks' as const, style, priority: 2 };
		const before = createWorkbook();
		before.sheets[0]!.conditionalFormats = [
			format(['A1', 'B1'], expression),
			format(['D1', 'E1'], blanks),
		];
		const after = createWorkbook();
		after.sheets[0]!.conditionalFormats = [
			format(['B1', 'A1'], expression),
			format(['E1', 'D1'], blanks),
		];
		expect(compareWorkbooks(before, after).sheets[0]!.structure.map((d) => d.kind)).toEqual([
			'conditionalFormatRemoved',
			'conditionalFormatAdded',
		]);
	});

	it('ignores data validation defaults spelled out', () => {
		const before = createWorkbook();
		const after = createWorkbook();
		before.sheets[0]!.dataValidations = [
			{ ranges: [R('A1:A9')], type: 'list', formula1: '"a,b"', showDropDown: true },
			{ ranges: [R('B1:B9')], type: 'whole', formula1: '1', formula2: '9' },
		];
		after.sheets[0]!.dataValidations = [
			{ ranges: [R('A1:A9')], type: 'list', formula1: '"a,b"', allowBlank: false, error: '' },
			{
				ranges: [R('B1:B9')],
				type: 'whole',
				operator: 'between',
				formula1: '1',
				formula2: '9',
				errorStyle: 'stop',
				showErrorMessage: false,
			},
		];
		expect(compareWorkbooks(before, after).count).toBe(0);
		// A hidden drop-down is a change.
		after.sheets[0]!.dataValidations[0]!.showDropDown = false;
		expect(compareWorkbooks(before, after).sheets[0]!.structure.map((d) => d.kind)).toEqual([
			'dataValidationRemoved',
			'dataValidationAdded',
		]);
	});

	it('counts structural differences against maxDifferences', () => {
		const before = book([['Data', 1]]);
		const after = book([
			['Data', 1],
			['More', 2],
		]);
		after.sheets[0]!.merges = [R('A1:B1'), R('A2:B2')];
		const result = compareWorkbooks(before, after, { maxDifferences: 2 });
		expect(result.workbook.map((d) => d.kind)).toEqual(['sheetAdded']);
		expect(result.sheets[0]!.structure.map((d) => d.kind)).toEqual(['mergeAdded']);
		expect([result.count, result.truncated]).toEqual([2, true]);
	});
});

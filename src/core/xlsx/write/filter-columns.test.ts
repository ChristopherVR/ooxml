import { describe, expect, it } from 'vitest';
import type { AutoFilterColumn } from '../model';
import { miniPackage, roundTrip, ws } from './mini-package-fixtures';

const ROWS = '<row r="1"><c r="A1"><v>1</v></c></row><row r="2"><c r="A2"><v>2</v></c></row>';
const STYLES =
	'<?xml version="1.0"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
	'<fonts count="1"><font><sz val="11"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>' +
	'<borders count="1"><border/></borders><cellStyleXfs count="1"><xf/></cellStyleXfs><cellXfs count="1"><xf/></cellXfs>' +
	'<dxfs count="2"><dxf><fill><patternFill><bgColor rgb="FFFF0000"/></patternFill></fill></dxf><dxf><font><color rgb="FF0000FF"/></font></dxf></dxfs></styleSheet>';

const pkg = (filter: string) =>
	miniPackage({ sheets: [{ name: 'S', xml: ws(ROWS, '', filter) }], styles: STYLES });

const column = async (xml: string): Promise<AutoFilterColumn> => {
	const { again } = await roundTrip(await pkg(`<autoFilter ref="A1:A2">${xml}</autoFilter>`));
	return again.sheets[0]!.autoFilter!.columns![0]!;
};

describe('autoFilter criteria model', () => {
	it('reads date groups with the checklist', async () => {
		const c = await column(
			'<filterColumn colId="0"><filters blank="1"><filter val="x"/><dateGroupItem year="2024" month="3" day="5" dateTimeGrouping="day"/></filters></filterColumn>',
		);
		expect(c.values).toEqual(['x']);
		expect(c.blank).toBe(true);
		expect(c.dateGroups).toEqual([{ year: 2024, month: 3, day: 5, grouping: 'day' }]);
	});

	it('reads custom, top 10, dynamic, colour and icon criteria', async () => {
		const custom = await column(
			'<filterColumn colId="0"><customFilters and="1"><customFilter operator="greaterThan" val="1"/><customFilter val="a*"/></customFilters></filterColumn>',
		);
		expect(custom.custom).toEqual({
			and: true,
			conditions: [
				{ operator: 'greaterThan', value: '1' },
				{ operator: 'equal', value: 'a*' },
			],
		});
		const top = await column(
			'<filterColumn colId="0"><top10 top="0" percent="1" val="25" filterVal="3"/></filterColumn>',
		);
		expect(top.top10).toEqual({ top: false, percent: true, value: 25, filterValue: 3 });
		const dynamic = await column(
			'<filterColumn colId="0"><dynamicFilter type="aboveAverage" val="2.5"/></filterColumn>',
		);
		expect(dynamic.dynamic).toEqual({ type: 'aboveAverage', value: 2.5 });
		const colour = await column(
			'<filterColumn colId="0"><colorFilter dxfId="1" cellColor="0"/></filterColumn>',
		);
		expect(colour.color?.cellColor).toBe(false);
		expect(colour.color?.dxfId).toBe(1);
		expect(colour.color?.style?.font?.color?.rgb).toBe('FF0000FF');
		const icon = await column(
			'<filterColumn colId="0"><iconFilter iconSet="3Arrows" iconId="2"/></filterColumn>',
		);
		expect(icon.icon).toEqual({ iconSet: '3Arrows', iconId: 2 });
	});

	it('writes edited criteria from the model and reads them back', async () => {
		const cases: Partial<AutoFilterColumn>[] = [
			{
				values: ['a', 'b'],
				blank: true,
				dateGroups: [{ year: 2020, month: 2, grouping: 'month' }],
			},
			{
				custom: {
					and: true,
					conditions: [
						{ operator: 'notEqual', value: 'x' },
						{ operator: 'equal', value: 'y' },
					],
				},
			},
			{ top10: { top: true, percent: false, value: 5 } },
			{ dynamic: { type: 'thisMonth' } },
			{ icon: { iconSet: '5Arrows', iconId: 0 } },
		];
		for (const extra of cases) {
			const { part, again } = await roundTrip(
				await pkg(
					'<autoFilter ref="A1:A2"><filterColumn colId="0"><filters><filter val="1"/></filters></filterColumn></autoFilter>',
				),
				(wb) => {
					wb.sheets[0]!.autoFilter!.columns = [{ offset: 0, ...extra }];
				},
			);
			const read = again.sheets[0]!.autoFilter!.columns![0]!;
			expect(read).toMatchObject(extra);
			expect(await part('xl/worksheets/sheet1.xml')).toContain('<filterColumn colId="0">');
		}
	});

	it('writes a new colour filter with a dxf from the style table', async () => {
		const { again } = await roundTrip(await pkg('<autoFilter ref="A1:A2"/>'), (wb) => {
			wb.sheets[0]!.autoFilter!.columns = [
				{
					offset: 0,
					color: {
						cellColor: true,
						style: { fill: { type: 'pattern', pattern: 'solid', fgColor: { rgb: 'FFFF0000' } } },
					},
				},
			];
		});
		const read = again.sheets[0]!.autoFilter!.columns![0]!;
		expect(read.color?.dxfId).toBe(0);
		expect(read.color?.style?.fill).toMatchObject({ pattern: 'solid' });
	});
});

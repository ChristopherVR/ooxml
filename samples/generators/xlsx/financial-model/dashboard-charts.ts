/** The sparklines and the two charts of the Dashboard, added through the edit session. */
import type { EditSession } from '../../../../src/core/xlsx/edit/types';
import { type Book, rgb } from '../lib/book';
import { anchor, chartObject } from '../lib/chart';
import { rangeOf } from '../lib/sheet';
import { FIRST, LAST, rowOf } from './dashboard';
import { P, SHEETS } from './style';

const S = SHEETS.dashboard;

/** Sparklines and charts, added through the session once the summary holds its values. */
export function decorateDashboard(book: Book, session: EditSession): void {
	const sheet = book.sheet(S);
	session.createSparklineGroup(
		sheet.index,
		`${S}!D${FIRST}:I${LAST}`,
		rangeOf(`J${FIRST}:J${LAST}`),
		{
			type: 'line',
			lineWeight: 1.5,
			high: true,
			low: true,
			colorSeries: rgb(P.walnut),
			colorHigh: rgb(P.sage),
			colorLow: rgb(P.berry),
			displayEmptyCellsAs: 'gap',
		},
	);
	const ref = (label: string, from = 'D') => {
		const row = rowOf(label);
		return {
			name: `${S}!$B$${row}`,
			categories: `${S}!$${from}$10:$I$10`,
			values: `${S}!$${from}$${row}:$I$${row}`,
		};
	};
	const top = LAST + 3;
	const bottom = top + 18;
	const common = { ink: P.ink, grid: P.rule, categoryFormat: '"FY"0', legend: 'b' as const };
	session.addChart(
		sheet.index,
		chartObject(
			book.wb,
			'Revenue and EBITDA margin',
			'Stacked columns of revenue by stream for FY2025 to FY2030 with the EBITDA margin as a line on a secondary axis.',
			anchor(`B${top}`, `G${bottom}`),
			{
				...common,
				title: 'Revenue by stream and EBITDA margin',
				valueFormat: '"$"#,##0,,"m"',
				secondaryFormat: '0%',
				secondaryMin: 0,
				secondaryMax: 0.3,
				groups: [
					{
						type: 'column',
						grouping: 'stacked',
						gapWidth: 55,
						series: [
							{ ...ref('Retail stores'), color: P.crust },
							{ ...ref('Wholesale'), color: P.honey },
							{ ...ref('Online orders'), color: P.teal },
						],
					},
					{
						type: 'line',
						secondary: true,
						series: [{ ...ref('EBITDA margin'), color: P.ink, markers: true, lineWidth: 2 }],
					},
				],
			},
		),
	);
	session.addChart(
		sheet.index,
		chartObject(
			book.wb,
			'Cash and bank debt',
			'Lines of year-end cash and bank debt for FY2025 to FY2030.',
			anchor(`G${top}`, `L${bottom}`),
			{
				...common,
				title: 'Cash builds while debt is repaid',
				valueFormat: '"$"#,##0,,"m"',
				valueMin: 0,
				groups: [
					{
						type: 'line',
						series: [
							{ ...ref('Cash at year end'), color: P.teal, markers: true },
							{ ...ref('Bank debt'), color: P.berry, markers: true },
						],
					},
				],
			},
		),
	);
}

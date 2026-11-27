/** The four charts and the two sparkline groups of the Dashboard, added through the edit session. */
import type { EditSession } from '../../../../src/core/xlsx/edit/types';
import { type Book, rgb } from '../lib/book';
import { anchor, chartObject } from '../lib/chart';
import { rangeOf } from '../lib/sheet';
import { AT } from './analysis';
import { REGION_FIRST, REGION_LAST } from './dashboard';
import { TERRITORIES } from './data';
import { P, SHEETS } from './style';

const S = SHEETS.dashboard;
const DATA = SHEETS.analysis;

/** Charts and sparklines, added through the session once the Analysis sheet is calculated. */
export function decorateDashboard(book: Book, session: EditSession): void {
	const sheet = book.sheet(S);
	const s = sheet.index;
	const col = (letter: string, first: number, last: number) =>
		`${DATA}!$${letter}$${first}:$${letter}$${last}`;
	const cell = (letter: string, row: number) => `${DATA}!$${letter}$${row}`;
	const monthsRef = col('B', AT.monthFirst, AT.monthLast);
	const monthSeries = (letter: string) => ({
		name: cell(letter, AT.monthHead),
		categories: monthsRef,
		values: col(letter, AT.monthFirst, AT.monthLast),
	});
	const common = { ink: P.ink, grid: P.rule };
	session.addChart(
		s,
		chartObject(
			book.wb,
			'Monthly revenue and margin',
			'Columns of monthly revenue for 2025 with the gross margin percentage as a line on a secondary axis.',
			anchor('B10', 'I27'),
			{
				...common,
				title: 'Monthly revenue and gross margin',
				legend: 'b',
				categoryFormat: 'mmm',
				valueFormat: '"$"#,##0,"k"',
				secondaryFormat: '0%',
				secondaryMin: 0.3,
				secondaryMax: 0.6,
				groups: [
					{ type: 'column', gapWidth: 45, series: [{ ...monthSeries('C'), color: P.pine }] },
					{
						type: 'line',
						secondary: true,
						series: [{ ...monthSeries('F'), color: P.ember, markers: true }],
					},
				],
			},
		),
	);
	session.addChart(
		s,
		chartObject(
			book.wb,
			'Channel mix',
			'A doughnut of revenue by sales channel.',
			anchor('I10', 'N27'),
			{
				...common,
				title: 'Revenue by channel',
				legend: 'b',
				groups: [
					{
						type: 'doughnut',
						holeSize: 58,
						series: [
							{
								name: cell('J', AT.categoryHead),
								categories: col('I', AT.channelFirst, AT.channelLast),
								values: col('J', AT.channelFirst, AT.channelLast),
								color: P.pine,
								pointColors: [P.pine, P.lake, P.ember],
								labels: '0%',
							},
						],
					},
				],
			},
		),
	);
	session.addChart(
		s,
		chartObject(
			book.wb,
			'Revenue by category',
			'Horizontal bars of revenue by product category, largest first.',
			anchor('B29', 'G46'),
			{
				...common,
				title: 'Revenue by category',
				valueFormat: '"$"#,##0,"k"',
				groups: [
					{
						type: 'bar',
						gapWidth: 50,
						series: [
							{
								name: cell('C', AT.categoryHead),
								categories: col('B', AT.categoryFirst, AT.categoryLast),
								values: col('C', AT.categoryFirst, AT.categoryLast),
								color: P.lake,
								labels: '"$"#,##0,"k"',
							},
						],
					},
				],
			},
		),
	);
	session.addChart(
		s,
		chartObject(
			book.wb,
			'Revenue by channel through the year',
			'Stacked areas of monthly revenue for the store, online and wholesale channels.',
			anchor('G29', 'N46'),
			{
				...common,
				title: 'Channels through the year',
				legend: 'b',
				categoryFormat: 'mmm',
				valueFormat: '"$"#,##0,"k"',
				groups: [
					{
						type: 'area',
						grouping: 'stacked',
						series: [
							{ ...monthSeries('K'), color: P.pine, opacity: 85 },
							{ ...monthSeries('L'), color: P.lake, opacity: 85 },
							{ ...monthSeries('M'), color: P.ember, opacity: 85 },
						],
					},
				],
			},
		),
	);
	const last = (first: number) => first + TERRITORIES.length - 1;
	session.createSparklineGroup(
		s,
		`${DATA}!C${AT.regionFirst}:N${last(AT.regionFirst)}`,
		rangeOf(`K${REGION_FIRST}:K${REGION_LAST}`),
		{
			type: 'line',
			lineWeight: 1.5,
			markers: true,
			high: true,
			colorSeries: rgb(P.pine),
			colorMarkers: rgb(P.pine),
			colorHigh: rgb(P.ember),
			displayEmptyCellsAs: 'gap',
		},
	);
	session.createSparklineGroup(
		s,
		`${DATA}!C${AT.countFirst}:N${last(AT.countFirst)}`,
		rangeOf(`L${REGION_FIRST}:L${REGION_LAST}`),
		{
			type: 'column',
			high: true,
			colorSeries: rgb(P.lake),
			colorHigh: rgb(P.ember),
			displayEmptyCellsAs: 'gap',
		},
	);
}

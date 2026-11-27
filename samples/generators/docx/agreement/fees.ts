/** Part C: the schedule of fees, a repeating section of priced lines under a calculated total. */
import type { Block, TableBorders } from '../../../../src/core/docx/index';
import { rich, t, type Doc } from '../kit/builder';
import { field } from '../kit/fields';
import { line, table, tw, type CellSpec } from '../kit/tables';
import { entry, formGrid, label, NO_BORDERS, partHeading, type Controls } from './controls';
import { INK, TEXT_WIDTH } from './styles';

/** The priced lines of the schedule: description, quantity and unit rate in euros. */
const LINES: [string, number, number][] = [
	['Brand discovery workshop and strategy note', 1, 1800],
	['Identity design: marque, palette and typefaces', 1, 2400],
	['Brand guidelines, 16 pages', 1, 650],
	['Stationery templates', 3, 150],
];
const money = (value: number): string => value.toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/gu, ',');
export const FEES_TOTAL = LINES.reduce((sum, [, quantity, rate]) => sum + quantity * rate, 0);

export function fees(doc: Doc, controls: Controls): Block[] {
	const right = { align: 'right' as const };
	const head = (text: string, extra: Partial<CellSpec> = {}): CellSpec => ({
		content: text,
		style: 'TableHead',
		fill: `#${INK.navy}`,
		...extra,
	});
	const rule: TableBorders = { bottom: line(INK.line, 0.5) };
	// Each priced line is one item of a repeating section: the "+" Word shows adds another row.
	const item = (index: number) =>
		controls.verbatim(
			'row',
			'repeatingSection',
			`Fee line ${index + 1}`,
			'fees.line',
			'<w15:repeatingSectionItem/>',
		);
	const lines = LINES.map(([description, quantity, rate], index) => [
		{
			content: controls.filled('Description', `fees.${index}.description`, description),
			borders: rule,
		},
		{
			content: controls.filled('Quantity', `fees.${index}.quantity`, String(quantity)),
			borders: rule,
			...right,
		},
		{
			content: controls.filled('Rate', `fees.${index}.rate`, money(rate)),
			borders: rule,
			...right,
		},
		{
			content: controls.filled('Amount', `fees.${index}.amount`, money(quantity * rate)),
			borders: rule,
			...right,
		},
	]);
	const total: TableBorders = { top: line(INK.navy, 1), bottom: line(INK.navy, 1) };
	return [
		partHeading(doc, 'C', 'Schedule of fees', {
			pageBreakBefore: true,
			spacingBeforeTwips: tw(120),
		}),
		doc.p(
			'Hint',
			'Add a line with the plus sign that appears at the end of a row. The total is a calculated field: select it and press F9 after a change.',
			{ keepNext: true },
		),
		table(doc, {
			style: 'FormGrid',
			textStyle: 'TableText',
			grid: [TEXT_WIDTH - 1000 - 1500 - 1700, 1000, 1500, 1700],
			borders: NO_BORDERS,
			look: {
				firstRow: true,
				lastRow: true,
				firstColumn: false,
				lastColumn: false,
				noHBand: true,
				noVBand: true,
			},
			cantSplit: true,
			rowProperties: [undefined, ...LINES.map((_, index) => ({ contentControl: item(index) }))],
			rows: [
				[
					head('Description'),
					head('Qty', right),
					head('Rate (€)', right),
					head('Amount (€)', right),
				],
				...lines,
				[
					{
						content: t('Total fees, excluding tax', { bold: true, color: `#${INK.navy}` }),
						span: 3,
						borders: total,
						...right,
					},
					{
						content: field(
							'=SUM(ABOVE) \\# "#,##0.00"',
							t(money(FEES_TOTAL), { bold: true, color: `#${INK.navy}` }),
						),
						borders: total,
						...right,
					},
				],
			],
		}),
		doc.p('Normal', '', {
			spacingAfterTwips: tw(0),
			lineSpacingTwips: tw(120),
			lineSpacingRule: 'exact',
		}),
		formGrid(doc, [
			[
				label('Payment terms'),
				entry(
					controls.choice(
						'comboBox',
						'Payment terms',
						'fees.terms',
						['On receipt', '14 days from invoice', '30 days from invoice'],
						'Choose, or type your own',
					),
				),
				label('Deposit'),
				entry(
					controls.choice(
						'dropDownList',
						'Deposit',
						'fees.deposit',
						['25% on signature', '40% on signature', '50% on signature'],
						'Choose a deposit',
					),
				),
			],
			[
				label('Order number'),
				entry([
					{
						text: '',
						fieldChar: 'begin',
						formField: {
							kind: 'text',
							name: 'OrderNumber',
							enabled: true,
							calcOnExit: false,
							statusText: { type: 'text', value: 'Your purchase order number, if you use one' },
							textInput: { default: 'PO-', maxLength: 16 },
						},
					},
					{ text: '', fieldCode: ' FORMTEXT ' },
					{ text: '', fieldChar: 'separate' },
					{ text: 'PO-', field: { instr: 'FORMTEXT' } },
					{ text: '', fieldChar: 'end' },
				]),
				label('Account manager'),
				entry(field('DOCPROPERTY "Account manager"', 'Amara Diallo')),
			],
		]),
		doc.p(
			'Hint',
			rich(
				'The order number is a legacy form field; like the check box in Part E it takes entries when the document is protected for filling in forms.',
			),
			{ spacingBeforeTwips: tw(100) },
		),
	];
}

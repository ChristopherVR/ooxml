/** The five busiest routes, as a formatted table beside a share card. */
import type { TableCellInput } from '../../../../src/core/pptx/core/builders/sdk/index';
import { FONTS } from '../../lib/fiction';
import type { Deck } from '../kit/deck';
import { TOP_ROUTES, TOP_SHARE, TOTAL_2025, change, signed, sum } from './data';
import { C, label, span } from './style';

/** A table without a style still draws every cell edge, so the unused edges are zeroed. */
const NO_LINE = { width: 0, color: C.paper };

function header(text: string, align: 'left' | 'right' = 'left'): TableCellInput {
	return {
		text,
		style: { fontSize: 9, bold: true, color: C.slate, fontFamily: FONTS.display, alignment: align },
		borders: { all: NO_LINE, bottom: { color: C.ink, width: 1.5 } },
	};
}

function cell(
	text: string,
	style: { bold?: boolean; color?: string; display?: boolean; align?: 'left' | 'right' } = {},
): TableCellInput {
	return {
		text,
		style: {
			fontSize: 12,
			bold: style.bold ?? false,
			color: style.color ?? C.ink,
			fontFamily: style.display ? FONTS.display : FONTS.body,
			alignment: style.align ?? 'left',
		},
		borders: { all: NO_LINE, bottom: { color: C.rule, width: 1 } },
	};
}

export function routesSlide(deck: Deck): void {
	const slide = deck.slide('Title only');
	const topTotal = sum(TOP_ROUTES.map((row) => row.boardings));
	slide.placeholder('title', `Five routes carry ${TOP_SHARE}% of all boardings, and all five grew`);
	slide.placeholder('body', 'The routes', 10);
	slide.placeholder('sldNum', '6', 12);
	const box = { ...span(0, 8), y: 208, height: 40 + TOP_ROUTES.length * 62 };
	const table = slide.table(
		'Top routes',
		{
			style: 'none',
			columnWidths: [64, 240, 92, 124, 112, 128],
			rows: [
				{
					height: 40,
					cells: [
						header('ROUTE'),
						header('LINE'),
						header('MODE'),
						header('BOARDINGS 2025', 'right'),
						header('ON 2024', 'right'),
						header('PEAK SERVICE', 'right'),
					],
				},
				...TOP_ROUTES.map((row, index) => ({
					height: 62,
					cells: [
						cell(row.route, { bold: true, display: true, color: index === 0 ? C.coral : C.teal }),
						cell(row.name, { bold: index === 0 }),
						cell(row.mode, { color: C.slate }),
						cell(`${row.boardings.toFixed(1)} million`, { align: 'right' }),
						cell(signed(change(row.boardings, row.before)), {
							bold: true,
							align: 'right',
							color: index === 0 ? C.coral : C.teal,
						}),
						cell(`every ${row.headway} min`, { color: C.slate, align: 'right' }),
					],
				})),
			],
		},
		box,
	);
	if (table.type === 'table' && table.tableData)
		for (const row of table.tableData.rows)
			for (const tableCell of row.cells)
				tableCell.style = {
					...tableCell.style,
					vAlign: 'middle',
					marginLeft: 10,
					marginRight: 10,
					marginTop: 4,
					marginBottom: 4,
				};
	const side = { ...span(9, 3), y: 208, height: box.height };
	const panel = slide.shape('Share card', 'roundRect', {
		...side,
		adjustments: { adj: 4200 },
		fill: {
			type: 'gradient',
			angle: 60,
			stops: [
				{ color: '#0B6A73', position: 0 },
				{ color: C.teal, position: 100 },
			],
		},
		stroke: { color: 'transparent' },
		shadow: { color: '#0B3A40', blur: 30, offsetX: 0, offsetY: 12, opacity: 0.22 },
	});
	const inner = { x: side.x + 24, width: side.width - 48 };
	const parts = [
		panel,
		label(
			slide,
			'Share figure',
			`${TOP_SHARE}%`,
			{ ...inner, y: side.y + 28, height: 70 },
			{
				size: 54,
				display: true,
				bold: true,
				color: C.white,
				tracking: -1.5,
				leading: 1,
			},
		),
		label(
			slide,
			'Share caption',
			`of all boardings: ${topTotal.toFixed(1)} of ${TOTAL_2025} million were made on these five routes.`,
			{ ...inner, y: side.y + 108, height: 80 },
			{ size: 12, color: '#DCEFF0', leading: 1.35 },
		),
		slide.shape('Share track', 'roundRect', {
			...inner,
			y: side.y + side.height - 96,
			height: 10,
			adjustments: { adj: 50000 },
			fill: { type: 'solid', color: '#3F959D' },
			stroke: { color: 'transparent' },
		}),
		slide.shape('Share bar', 'roundRect', {
			x: inner.x,
			y: side.y + side.height - 96,
			width: (inner.width * TOP_SHARE) / 100,
			height: 10,
			adjustments: { adj: 50000 },
			fill: { type: 'solid', color: C.coral },
			stroke: { color: 'transparent' },
		}),
		label(
			slide,
			'Share note',
			'Two tram lines, two bus routes and one ferry.',
			{ ...inner, y: side.y + side.height - 72, height: 44 },
			{ size: 10.5, color: '#DCEFF0', leading: 1.3 },
		),
	];
	label(
		slide,
		'Table note',
		'Boardings in 2025. Peak service is the weekday frequency between 7 and 9 am.',
		{ ...span(0, 8), y: box.y + box.height + 20, height: 18 },
		{ size: 9.5, color: C.slate },
	);
	slide.animate(table, { entrance: 'fadeIn', trigger: 'afterPrevious', durationMs: 600 });
	parts.forEach((part, index) =>
		slide.animate(part, {
			entrance: index === 4 ? 'wipeIn' : 'fadeIn',
			...(index === 4 ? { direction: 'fromLeft' as const } : {}),
			trigger: index === 0 ? 'onClick' : index === 4 ? 'afterPrevious' : 'withPrevious',
			durationMs: 500,
		}),
	);
	slide.transition({ type: 'split', direction: 'out', duration: 700 });
	slide.notes(
		`The five busiest routes carried ${topTotal.toFixed(1)} million boardings, ${TOP_SHARE}% of the network total. T1, the Harbour Line, grew ${signed(change(TOP_ROUTES[0]!.boardings, TOP_ROUTES[0]!.before))} and is now our busiest route by a wide margin. Route 4 is the first Frequent Network bus route to pass three million boardings. Click to bring in the share card.`,
	);
}

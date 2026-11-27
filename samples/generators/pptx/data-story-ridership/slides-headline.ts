/** Section "The headline": the cover, the year in one number, and the year month by month. */
import type { Deck } from '../kit/deck';
import { monthlyChart } from './charts';
import {
	MONTHS_AHEAD,
	PEAK_2019,
	RECORD_MONTH,
	TOTAL_2019,
	TOTAL_2024,
	TOTAL_2025,
	change,
	signed,
} from './data';
import { annotation, legend, stat } from './parts';
import { C, CONTENT_TOP, MARGIN, frame, label, span } from './style';

const SOURCE =
	'Source: Port Averly Transit automatic passenger counts, validated monthly totals. Boardings in millions.';

export function headlineSlides(deck: Deck): void {
	deck.section('The headline');
	const onLastYear = signed(change(TOTAL_2025, TOTAL_2024));
	const onRecord = signed(change(TOTAL_2025, TOTAL_2019));
	const overPeak = Math.round(change(RECORD_MONTH, PEAK_2019));

	const cover = deck.slide('Cover');
	cover.placeholder('ctrTitle', 'How Port Averly moved');
	cover.placeholder(
		'subTitle',
		`${TOTAL_2025} million boardings in 2025: where they came from, and what we do next.`,
	);
	cover.placeholder('body', 'Ridership review 2025', 10);
	cover.placeholder('body', 'Board briefing, 15 January 2026', 11);
	cover.transition({ type: 'fade', duration: 700 });
	cover.notes(
		`Good morning. This is the 2025 ridership review. The short version: ${TOTAL_2025} million boardings, a record for the network. The next twelve slides show where the growth came from, why we think it happened, and what we propose for 2026.`,
	);

	const year = deck.slide('Statement');
	frame(
		year,
		{
			kicker: 'The headline',
			headline: 'Port Averly made more trips by transit in 2025 than in any year on record',
			source: SOURCE,
		},
		2,
	);
	const total = label(
		year,
		'!!total',
		String(TOTAL_2025),
		{ x: MARGIN - 6, y: 208, width: 420, height: 170 },
		{ size: 132, display: true, bold: true, color: C.coral, tracking: -4, leading: 1 },
	);
	const bar = year.shape('!!bar', 'rect', {
		x: MARGIN,
		y: 392,
		width: 400,
		height: 8,
		fill: {
			type: 'gradient',
			angle: 0,
			stops: [
				{ color: C.coral, position: 0 },
				{ color: C.ochre, position: 100 },
			],
		},
		stroke: { color: 'transparent' },
	});
	const unit = label(
		year,
		'!!unit',
		'million boardings in 2025',
		{ x: 484, y: 318, width: 420, height: 40 },
		{ size: 22, display: true, color: C.ink },
	);
	const columns = [0, 4, 8].map((start) => ({
		...span(start, 4),
		width: span(start, 4).width - 40,
	}));
	const stats = [
		stat(year, 'Stat 2024', onLastYear, `on 2024, when the network carried ${TOTAL_2024} million`, {
			...columns[0]!,
			y: 456,
			height: 130,
		}),
		stat(
			year,
			'Stat 2019',
			onRecord,
			`on 2019, the previous record year at ${TOTAL_2019} million`,
			{
				...columns[1]!,
				y: 456,
				height: 130,
			},
		),
		stat(
			year,
			'Stat October',
			`${RECORD_MONTH}M`,
			'boardings in October, the busiest month the network has had',
			{ ...columns[2]!, y: 456, height: 130 },
		),
	];
	year.animate(total, { entrance: 'zoomIn', trigger: 'afterPrevious', durationMs: 600 });
	year.animate(unit, { entrance: 'fadeIn', trigger: 'withPrevious', durationMs: 600 });
	year.animate(bar, {
		entrance: 'wipeIn',
		direction: 'fromLeft',
		trigger: 'afterPrevious',
		durationMs: 500,
	});
	year.animate(total, { emphasis: 'pulse', trigger: 'afterPrevious', durationMs: 700 });
	stats.forEach((parts, index) =>
		parts.forEach((part, at) =>
			year.animate(part, {
				entrance: 'fadeIn',
				trigger: at === 0 ? (index === 0 ? 'onClick' : 'afterPrevious') : 'withPrevious',
				durationMs: 450,
			}),
		),
	);
	year.transition({ type: 'push', direction: 'l', duration: 600 });
	year.notes(
		`${TOTAL_2025} million boardings is ${onLastYear} on 2024 and ${onRecord} on 2019, which was our record until now. Click once to bring in the three supporting figures. The October figure matters most: ${RECORD_MONTH} million is the busiest month this network has ever had.`,
	);

	const months = deck.slide('Chart');
	frame(
		months,
		{
			kicker: 'The year, month by month',
			headline: `Every month since April beat 2019, and October set an all-time record`,
			source: SOURCE,
		},
		3,
	);
	const side = span(9, 3);
	months.shape('!!bar', 'rect', {
		x: side.x,
		y: 72,
		width: side.width,
		height: 3,
		fill: {
			type: 'gradient',
			angle: 0,
			stops: [
				{ color: C.coral, position: 0 },
				{ color: C.ochre, position: 100 },
			],
		},
		stroke: { color: 'transparent' },
	});
	label(
		months,
		'!!total',
		String(TOTAL_2025),
		{ x: side.x, y: 84, width: 140, height: 56 },
		{ size: 40, display: true, bold: true, color: C.coral, tracking: -1, leading: 1 },
	);
	label(
		months,
		'!!unit',
		'million boardings in 2025',
		{ x: side.x + 132, y: 94, width: side.width - 132, height: 40 },
		{ size: 11, color: C.slate, leading: 1.25 },
	);
	legend(
		months,
		[
			{ label: '2019', color: C.sandDeep },
			{ label: '2025', color: C.teal },
			{ label: 'Record month', color: C.coral },
		],
		MARGIN,
		CONTENT_TOP,
	);
	const chart = months.chart('Monthly boardings', 'bar', monthlyChart(), {
		x: MARGIN - 8,
		y: 228,
		width: 1168,
		height: 388,
	});
	const note = annotation(
		months,
		'October',
		`October: ${RECORD_MONTH} million, ${overPeak}% above the 2019 peak of ${PEAK_2019} million`,
		{ x: 604, y: 196, width: 330, height: 34 },
		{ x: 984, y: 262 },
	);
	months.animate(chart, {
		entrance: 'wipeIn',
		direction: 'fromBottom',
		trigger: 'afterPrevious',
		durationMs: 900,
	});
	note.forEach((part, index) =>
		months.animate(part, {
			entrance: 'fadeIn',
			trigger: index === 0 ? 'afterPrevious' : 'withPrevious',
			durationMs: 500,
		}),
	);
	months.transition({ type: 'morph', duration: 1100 });
	months.notes(
		`The total becomes the scale for this chart. Sand columns are 2019, teal are 2025. We started the year behind 2019, drew level in April and stayed ahead for ${MONTHS_AHEAD} months. October, in coral, reached ${RECORD_MONTH} million: ${overPeak}% above the old peak of ${PEAK_2019} million, which was also an October.`,
	);
}

/** Section "Why", second half: the evidence for each decision (fares, service, frequency). */
import type { Deck } from '../kit/deck';
import { frequencyChart, productivityChart, quietSecondAxis } from './charts';
import {
	FREQUENCY_RATIO,
	FREQUENT_MEAN,
	FREQUENT_ROUTES,
	HOURS_2024,
	HOURS_2025,
	OTHER_MEAN,
	PRODUCTIVITY_2024,
	PRODUCTIVITY_2025,
	TOTAL_2024,
	TOTAL_2025,
	change,
	signed,
	thousands,
} from './data';
import { annotation, legend, stat } from './parts';
import { faresSlide } from './slides-fares';
import { C, CONTENT_TOP, MARGIN, frame, label, span } from './style';

function serviceSlide(deck: Deck): void {
	const slide = deck.slide('Chart');
	const hours = signed(change(HOURS_2025, HOURS_2024));
	const boardings = signed(change(TOTAL_2025, TOTAL_2024));
	frame(
		slide,
		{
			kicker: 'Why: service',
			headline: `We ran ${hours.slice(1)} more service and carried ${boardings.slice(1)} more people`,
			source:
				'Revenue service hours (thousands, left axis) and boardings per service hour (right axis), by quarter. Source: Port Averly Transit.',
		},
		10,
	);
	legend(
		slide,
		[
			{ label: 'Service hours, thousands', color: C.sandDeep },
			{ label: 'Boardings per service hour', color: C.coral },
		],
		MARGIN,
		CONTENT_TOP,
	);
	const chart = slide.chart('Service and productivity', 'bar', productivityChart(), {
		x: MARGIN - 8,
		y: 228,
		width: span(0, 9).width + 16,
		height: 388,
	});
	quietSecondAxis(chart, { min: 20, max: 40, majorUnit: 5 });
	const side = span(9, 3);
	const stats = [
		stat(
			slide,
			'Stat hours',
			hours,
			`service hours: ${thousands(HOURS_2024)} in 2024, ${thousands(HOURS_2025)} in 2025`,
			{
				...side,
				y: 196,
				height: 130,
			},
		),
		stat(
			slide,
			'Stat boardings',
			boardings,
			`boardings: ${TOTAL_2024} million in 2024, ${TOTAL_2025} million in 2025`,
			{
				...side,
				y: 344,
				height: 130,
			},
		),
		stat(
			slide,
			'Stat productivity',
			String(PRODUCTIVITY_2025),
			`boardings per service hour in 2025, up from ${PRODUCTIVITY_2024}`,
			{ ...side, y: 492, height: 120 },
			C.coral,
		),
	];
	slide.animate(chart, {
		entrance: 'wipeIn',
		direction: 'fromLeft',
		trigger: 'afterPrevious',
		durationMs: 900,
	});
	stats.forEach((parts, index) =>
		parts.forEach((part, at) =>
			slide.animate(part, {
				entrance: 'fadeIn',
				trigger: at > 0 ? 'withPrevious' : index === 0 ? 'onClick' : 'afterPrevious',
				durationMs: 400,
			}),
		),
	);
	slide.transition({ type: 'wipe', direction: 'd', duration: 700 });
	slide.notes(
		`This is the efficiency slide. The sand columns are service hours: ${thousands(HOURS_2024)} in 2024 and ${thousands(HOURS_2025)} in 2025, ${hours}. The coral line is boardings per service hour, which rose from ${PRODUCTIVITY_2024} to ${PRODUCTIVITY_2025} across the year. We did not buy this growth with service alone; each hour is carrying more people.`,
	);
}

function frequencySlide(deck: Deck): void {
	const slide = deck.slide('Chart');
	const ratio = FREQUENCY_RATIO.toFixed(1);
	frame(
		slide,
		{
			kicker: 'Why: frequency',
			headline: `Frequency pays: Frequent Network routes carry ${ratio} times the boardings per service hour`,
			source:
				'All 24 bus routes, 2025. Each dot is one route. Source: Port Averly Transit schedule and passenger counts.',
		},
		11,
	);
	legend(
		slide,
		[
			{ label: 'Frequent Network, 6 routes', color: C.coral },
			{ label: 'Other bus routes, 18', color: C.tealSoft },
		],
		MARGIN,
		CONTENT_TOP,
	);
	const chart = slide.chart('Frequency and productivity', 'scatter', frequencyChart(), {
		x: MARGIN - 8,
		y: 250,
		width: span(0, 9).width + 16,
		height: 350,
	});
	quietSecondAxis(chart, { min: 0, max: 70, majorUnit: 10 });
	label(
		slide,
		'Axis x',
		'Buses per hour in the peak',
		{ x: MARGIN + 30, y: 598, width: span(0, 9).width - 30, height: 16 },
		{ size: 9.5, color: C.slate, align: 'center' },
	);
	label(
		slide,
		'Axis y',
		'Boardings per service hour',
		{ x: MARGIN, y: 232, width: 300, height: 16 },
		{ size: 9.5, color: C.slate },
	);
	// The plot area of this chart box runs from 0 at x 94 to 14 at x 908, and from 0 at y 563
	// to 70 at y 264; the leader ends just above the busiest Frequent Network route.
	const best = FREQUENT_ROUTES.reduce((a, b) => (b.productivity > a.productivity ? b : a));
	const note = annotation(
		slide,
		'Frequent',
		'A bus every 8 minutes or better',
		{ x: 470, y: 270, width: 220, height: 20 },
		{ x: 94 + best.frequency * 58.14, y: 563 - best.productivity * 4.27 - 14 },
	);
	const side = span(9, 3);
	const stats = [
		stat(
			slide,
			'Stat frequent',
			FREQUENT_MEAN.toFixed(1),
			'boardings per service hour on the six Frequent Network routes',
			{
				...side,
				y: 196,
				height: 130,
			},
			C.coral,
		),
		stat(
			slide,
			'Stat other',
			OTHER_MEAN.toFixed(1),
			'on the other 18 bus routes, which run two to six buses an hour',
			{
				...side,
				y: 344,
				height: 130,
			},
		),
	];
	label(
		slide,
		'Frequency reading',
		'More buses per hour do not just add capacity. They remove the timetable, and riders respond.',
		{ ...side, y: 492, height: 110 },
		{ size: 13, color: C.ink, leading: 1.35 },
	);
	slide.animate(chart, { entrance: 'fadeIn', trigger: 'afterPrevious', durationMs: 700 });
	note.forEach((part, index) =>
		slide.animate(part, {
			entrance: 'fadeIn',
			trigger: index === 0 ? 'afterPrevious' : 'withPrevious',
			durationMs: 400,
		}),
	);
	stats.flat().forEach((part, index) =>
		slide.animate(part, {
			entrance: 'fadeIn',
			trigger: index === 0 ? 'onClick' : 'withPrevious',
			durationMs: 400,
		}),
	);
	slide.transition({ type: 'push', direction: 'l', duration: 600 });
	slide.notes(
		`Every dot is one of our 24 bus routes. Across the chart, more buses per hour means more boardings for every hour of service. The six coral dots are the Frequent Network: they average ${FREQUENT_MEAN.toFixed(1)} boardings per service hour, against ${OTHER_MEAN.toFixed(1)} for the other 18 routes, a ratio of ${ratio}. That is the case for extending the network in 2026.`,
	);
}

export function evidenceSlides(deck: Deck): void {
	faresSlide(deck);
	serviceSlide(deck);
	frequencySlide(deck);
}

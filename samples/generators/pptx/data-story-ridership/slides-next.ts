/** Section "What next": the 2026 plan as SmartArt, and the closing takeaways. */
import type { PptxSmartArtData } from '../../../../src/core/pptx/core/types/smart-art';
import { applyBuiltinSmartArtLayout } from '../../../../src/core/pptx/smartart-layouts/index';
import { PEOPLE } from '../../lib/fiction';
import type { Deck } from '../kit/deck';
import {
	PRODUCTIVITY_2024,
	PRODUCTIVITY_2025,
	TARGET_2026,
	TOTAL_2019,
	TOTAL_2025,
	change,
	signed,
	sum,
} from './data';
import { C, MARGIN, label, rule, span } from './style';

/** The four steps of the plan and the boardings each is expected to add, in millions. */
const PLAN = [
	{
		quarter: 'Q1',
		title: 'Night trams',
		detail: 'T1 and T2 run until 2 am on Fridays and Saturdays.',
		gain: 0.5,
	},
	{
		quarter: 'Q2',
		title: '9 frequent routes',
		detail: 'Routes 2, 7 and 11 move to a bus every 8 minutes.',
		gain: 1.4,
	},
	{
		quarter: 'Q3',
		title: 'Greywharf ferry',
		detail: 'A second ferry route, F2, with a boat every 20 minutes.',
		gain: 0.4,
	},
	{
		quarter: 'Q4',
		title: 'All-door boarding',
		detail: 'Tap readers at every door of every bus, to cut time at stops.',
		gain: 0.3,
	},
];
/** What a full year of the 2025 changes adds on its own. */
const CARRY_OVER = 0.8;

function planSlide(deck: Deck): void {
	const slide = deck.slide('Title only');
	const growth = signed(change(TARGET_2026, TOTAL_2025));
	const planned = sum([...PLAN.map((step) => step.gain), CARRY_OVER]);
	if (sum([TOTAL_2025, planned]) !== TARGET_2026)
		throw new Error('The 2026 plan does not add up to its target.');
	slide.placeholder(
		'title',
		`2026: extend what worked, and aim for ${TARGET_2026} million boardings`,
	);
	slide.placeholder('body', 'What next', 10);
	slide.placeholder('sldNum', '12', 12);
	const steps: PptxSmartArtData = {
		layout: 'basicChevronProcess',
		colorScheme: 'monochromatic1',
		style: 'flat',
		nodes: PLAN.map((step, index) => ({ id: `plan${index + 1}`, text: step.title })),
	};
	const data = applyBuiltinSmartArtLayout(
		steps,
		'urn:microsoft.com/office/officeart/2005/8/layout/chevron1',
	);
	const diagram = slide.smartArt('Plan for 2026', data, {
		x: MARGIN,
		y: 204,
		width: span(0, 12).width,
		height: 132,
	});
	const details = PLAN.flatMap((step, index) => {
		const column = { ...span(index * 3, 3), x: span(index * 3, 3).x + 8 };
		const width = column.width - 24;
		return [
			label(
				slide,
				`Plan quarter ${index + 1}`,
				`${step.quarter} 2026`,
				{ x: column.x, y: 360, width, height: 16 },
				{
					size: 9.5,
					display: true,
					bold: true,
					color: C.slate,
					tracking: 1.6,
				},
			),
			label(
				slide,
				`Plan gain ${index + 1}`,
				`+${step.gain.toFixed(1)}M`,
				{ x: column.x, y: 382, width, height: 40 },
				{
					size: 26,
					display: true,
					bold: true,
					color: C.teal,
					tracking: -0.5,
				},
			),
			label(
				slide,
				`Plan detail ${index + 1}`,
				step.detail,
				{ x: column.x, y: 426, width, height: 60 },
				{
					size: 12,
					color: C.slate,
					leading: 1.35,
				},
			),
		];
	});
	rule(slide, 'Target rule', { x: MARGIN, y: 508, width: span(0, 12).width, height: 2 }, C.ink);
	const target = [
		label(
			slide,
			'Target figure',
			`${TARGET_2026}.0 million`,
			{ x: MARGIN, y: 528, width: 420, height: 56 },
			{
				size: 40,
				display: true,
				bold: true,
				color: C.coral,
				tracking: -1,
				leading: 1,
			},
		),
		label(
			slide,
			'Target caption',
			`boardings in 2026, ${growth} on 2025. The four steps add ${sum(PLAN.map((step) => step.gain)).toFixed(1)} million; a full year of the 2025 changes adds the other ${CARRY_OVER.toFixed(1)} million.`,
			{ ...span(4, 8), y: 534, height: 60 },
			{ size: 13, color: C.ink, leading: 1.4 },
		),
	];
	slide.animate(diagram, {
		entrance: 'wipeIn',
		direction: 'fromLeft',
		trigger: 'afterPrevious',
		durationMs: 900,
	});
	details.forEach((part, index) =>
		slide.animate(part, {
			entrance: 'fadeIn',
			trigger: index === 0 ? 'afterPrevious' : 'withPrevious',
			durationMs: 500,
			delayMs: Math.floor(index / 3) * 150,
		}),
	);
	target.forEach((part, index) =>
		slide.animate(part, {
			entrance: 'fadeIn',
			trigger: index === 0 ? 'onClick' : 'withPrevious',
			durationMs: 500,
		}),
	);
	slide.transition({ type: 'zoom', direction: 'in', duration: 700 });
	slide.notes(
		`The plan for 2026 is more of what worked. One step each quarter: night trams, three more Frequent Network routes, a second ferry and all-door boarding. Together they are forecast to add ${sum(PLAN.map((step) => step.gain)).toFixed(1)} million boardings. With ${CARRY_OVER.toFixed(1)} million from a full year of the 2025 changes, the target is ${TARGET_2026} million, ${growth} on this year.`,
	);
}

function closingSlide(deck: Deck): void {
	const slide = deck.slide('Closing');
	slide.placeholder('title', 'What we learned in 2025');
	slide.placeholder('body', 'In short', 10);
	slide.placeholder(
		'body',
		[
			`Ridership is at a record: ${TOTAL_2025} million boardings, ${signed(change(TOTAL_2025, TOTAL_2019))} on 2019.`,
			'Growth followed investment: new track, frequent buses, simpler fares.',
			`Service is working harder: ${PRODUCTIVITY_2025} boardings per service hour, up from ${PRODUCTIVITY_2024}.`,
			`Next: extend what worked, and aim for ${TARGET_2026} million in 2026.`,
		],
		13,
	);
	slide.placeholder('sldNum', '13', 12);
	const side = span(8, 4);
	slide.shape('Contact rule', 'rect', {
		x: side.x,
		y: 204,
		width: 56,
		height: 4,
		fill: { type: 'solid', color: C.coral },
		stroke: { color: 'transparent' },
	});
	label(
		slide,
		'Contact kicker',
		'QUESTIONS',
		{ x: side.x, y: 228, width: side.width, height: 16 },
		{
			size: 9.5,
			display: true,
			bold: true,
			color: C.tealSoft,
			tracking: 1.6,
		},
	);
	label(
		slide,
		'Contact name',
		PEOPLE[2],
		{ x: side.x, y: 254, width: side.width, height: 30 },
		{
			size: 20,
			display: true,
			bold: true,
			color: C.paper,
		},
	);
	label(
		slide,
		'Contact details',
		['Director of Network Planning', 'Port Averly Transit', 'planning@portaverlytransit.example'],
		{ x: side.x, y: 290, width: side.width, height: 80 },
		{ size: 12, color: '#A9B3C1', leading: 1.5 },
	);
	// The three figures the deck rests on, as a closing line.
	const recap = [
		{ figure: `${TOTAL_2025}M`, caption: 'boardings in 2025', color: C.coral },
		{ figure: String(PRODUCTIVITY_2025), caption: 'boardings per service hour', color: C.paper },
		{ figure: `${TARGET_2026}.0M`, caption: 'target for 2026', color: C.paper },
	];
	recap.forEach((item, index) => {
		const column = span(index * 3, 3);
		slide.shape(`Recap rule ${index + 1}`, 'rect', {
			x: column.x,
			y: 542,
			width: column.width - 24,
			height: 1,
			fill: { type: 'solid', color: '#3A4656' },
			stroke: { color: 'transparent' },
		});
		label(
			slide,
			`Recap figure ${index + 1}`,
			item.figure,
			{ x: column.x, y: 554, width: column.width, height: 44 },
			{
				size: 30,
				display: true,
				bold: true,
				color: item.color,
				tracking: -0.5,
			},
		);
		label(
			slide,
			`Recap caption ${index + 1}`,
			item.caption,
			{ x: column.x, y: 608, width: column.width, height: 18 },
			{
				size: 11,
				color: '#A9B3C1',
			},
		);
	});
	slide.effects.push(
		{ target: 'Body 3', paragraph: 0, kind: 'rise', start: 'after' },
		{ target: 'Body 3', paragraph: 1, kind: 'rise', start: 'click' },
		{ target: 'Body 3', paragraph: 2, kind: 'rise', start: 'click' },
		{ target: 'Body 3', paragraph: 3, kind: 'rise', start: 'click' },
	);
	slide.transition({ type: 'fade', duration: 800 });
	slide.notes(
		`Four things to take away. The first appears with the slide, the other three on a click each. Ridership is at a record, the growth followed specific investments, the service is more productive than it was, and the 2026 plan extends the same three ideas. ${PEOPLE[2]} and the planning team are happy to take questions.`,
	);
}

export function nextSlides(deck: Deck): void {
	deck.section('What next');
	planSlide(deck);
	closingSlide(deck);
}

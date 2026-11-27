/** Slides 11 to 13: price and availability, the pre-orders (a doughnut chart) and the close. */
import { REGIONS } from '../../lib/fiction';
import type { Deck } from '../kit/deck';
import { card, head, tx } from './parts';
import { C, G, T, body, col, display, label, noLine } from './style';
import { LAYOUT } from './template';

const TIERS = [
	{
		name: 'Aerie',
		price: '$249',
		text: 'The station, the mast kit and the app. Eleven signals, solar powered.',
		foot: 'Launch price until 2 March',
	},
	{
		name: 'Aerie Pro',
		price: '$329',
		text: 'Adds a soil and leaf-wetness probe and a LoRa bridge with a 2 km range.',
		foot: 'For gardens, vineyards and smallholdings',
	},
	{
		name: 'Aerie Cloud',
		price: 'Free',
		text: 'Your forecast and every alert, at no charge. Plus adds ten years of history for $4 a month.',
		foot: 'Included with every station',
	},
] as const;

function pricing(deck: Deck): void {
	const slide = deck.slide(LAYOUT.titleOnly);
	head(slide, 'Price and availability', 'From $249. On roofs from 3 March.');
	const top = 212;
	TIERS.forEach((tier, index) => {
		const { x, width } = col(index * 4, 4);
		const inner = { x: x + 28, width: width - 56 };
		card(slide, `Tier ${index + 1}`, { x, y: top, width, height: 316 });
		tx(
			slide,
			`Tier ${index + 1} name`,
			tier.name.toUpperCase(),
			{ ...inner, y: top + 30, height: 18 },
			label(index === 0 ? C.amber : C.sky),
		);
		tx(
			slide,
			`Tier ${index + 1} price`,
			tier.price,
			{ ...inner, y: top + 62, height: 72 },
			display(T.stat, index === 0 ? C.amber : C.text),
		);
		tx(
			slide,
			`Tier ${index + 1} text`,
			tier.text,
			{ ...inner, y: top + 150, height: 88 },
			body(T.small + 1),
			{ lineSpacing: 1.3 },
		);
		slide.shape(`Tier ${index + 1} rule`, 'rect', {
			...inner,
			y: top + 258,
			height: 1,
			fill: { type: 'solid', color: C.line },
			stroke: noLine,
		});
		tx(
			slide,
			`Tier ${index + 1} foot`,
			tier.foot,
			{ ...inner, y: top + 272, height: 20 },
			body(T.caption, C.muted),
		);
	});
	// The struck-through list price sits beside the launch price and leaves on the first click.
	const first = col(0, 4);
	const list = tx(
		slide,
		'List price',
		[[{ text: '$289', style: { strikethrough: true } }]],
		{ x: first.x + 214, y: top + 92, width: 110, height: 32 },
		display(T.lead, C.muted),
	);
	const pill = slide.shape('Pre-orders open', 'roundRect', {
		x: G.left,
		y: 564,
		width: 250,
		height: 44,
		fill: {
			type: 'gradient',
			angle: 0,
			stops: [
				{ color: C.amber, position: 0 },
				{ color: C.coral, position: 100 },
			],
		},
		stroke: noLine,
		adjustments: { adj: 50000 },
		textStyle: { alignment: 'center', verticalAlignment: 'middle' },
		text: [{ text: 'Pre-orders open today', style: display(12.5, C.night) }],
	});
	const ships = tx(
		slide,
		'Shipping',
		'Shipping from 3 March 2026 to all five regions.',
		{ x: G.left + 274, y: 564, width: 600, height: 44 },
		body(T.body),
		{ verticalAlignment: 'middle' },
	);
	slide
		.animate(list, { exit: 'fadeOut', trigger: 'onClick', durationMs: 600 })
		.animate(pill, {
			entrance: 'flyIn',
			direction: 'fromLeft',
			trigger: 'onClick',
			durationMs: 500,
		})
		.animate(ships, { entrance: 'fadeIn', trigger: 'withPrevious', durationMs: 500, delayMs: 200 })
		.transition({ type: 'wipe', direction: 'u', duration: 700 })
		.notes(
			'The first card shows the launch price of 249 next to the list price of 289. Click once and the list price fades out: 249 holds until 2 March. Aerie Pro is 329, and the cloud service is free for your own station. Click again for the pre-order banner: shipping starts on 3 March in all five regions.',
		);
}

/** Pre-orders by region; they add up to 12,000. */
const ORDERS: readonly [region: string, orders: number, color: string][] = [
	[REGIONS[0], 4320, C.amber],
	[REGIONS[3], 2640, C.sky],
	[REGIONS[1], 2280, C.violet],
	[REGIONS[4], 1560, '#5EE0B5'],
	[REGIONS[2], 1200, '#56668F'],
];

function preOrders(deck: Deck): void {
	const slide = deck.slide(LAYOUT.titleOnly);
	const total = ORDERS.reduce((sum, [, orders]) => sum + orders, 0);
	const count = (value: number) => String(value).replace(/\B(?=(\d{3})+$)/gu, ',');
	head(slide, 'Pre-orders', `${count(total)} roofs are already waiting`);
	const chart = slide.chart(
		'Pre-orders by region',
		'doughnut',
		{
			categories: ORDERS.map(([region]) => region),
			series: [
				{
					name: 'Pre-orders',
					values: ORDERS.map(([, orders]) => orders),
					pointColors: ORDERS.map(([, , color]) => color),
				},
			],
			embedWorkbook: true,
			hasLegend: false,
			holeSize: 70,
			firstSliceAngle: 0,
		},
		{ x: G.left, y: 196, width: 440, height: 440 },
	);
	tx(slide, 'Total', count(total), { x: G.left, y: 372, width: 440, height: 60 }, display(40), {
		alignment: 'center',
	});
	tx(slide, 'Total label', 'PRE-ORDERS', { x: G.left, y: 434, width: 440, height: 18 }, label(), {
		alignment: 'center',
	});
	const left = 620;
	const span = G.right - left;
	const bars = ORDERS.map(([region, orders, color], index) => {
		const y = 214 + index * 82;
		slide.shape(`Region ${index + 1} dot`, 'ellipse', {
			x: left,
			y: y + 6,
			width: 12,
			height: 12,
			fill: { type: 'solid', color },
			stroke: noLine,
		});
		tx(
			slide,
			`Region ${index + 1}`,
			region,
			{ x: left + 28, y, width: 300, height: 26 },
			display(15),
		);
		tx(
			slide,
			`Region ${index + 1} orders`,
			[
				[
					{ text: count(orders) },
					{ text: `   ${Math.round((orders / total) * 100)} %`, style: body(T.small, C.muted) },
				],
			],
			{ x: left + span - 220, y, width: 220, height: 26 },
			display(15),
			{ alignment: 'right' },
		);
		slide.shape(`Region ${index + 1} track`, 'rect', {
			x: left + 28,
			y: y + 40,
			width: span - 28,
			height: 4,
			fill: { type: 'solid', color: C.panelHi },
			stroke: noLine,
		});
		return slide.shape(`Region ${index + 1} bar`, 'rect', {
			x: left + 28,
			y: y + 40,
			width: ((span - 28) * orders) / ORDERS[0]![1],
			height: 4,
			fill: { type: 'solid', color },
			stroke: noLine,
		});
	});
	slide.animate(chart, { entrance: 'wheelIn', trigger: 'afterPrevious', durationMs: 1000 });
	bars.forEach((bar, index) =>
		slide.animate(bar, {
			entrance: 'wipeIn',
			direction: 'fromLeft',
			trigger: index === 0 ? 'afterPrevious' : 'withPrevious',
			durationMs: 600,
			delayMs: index * 120,
		}),
	);
	slide
		.transition({ type: 'reveal', duration: 900 })
		.notes(
			'We opened a quiet waiting list to the beta community in November. Twelve thousand roofs are on it: 4,320 on the Northwind Coast, where the weather comes in sideways, then Lakeshore, the Central Plains, the Southern Delta and Highland. The chart is a native doughnut with its workbook embedded.',
		);
}

function close(deck: Deck): void {
	const slide = deck.slide(LAYOUT.closing);
	slide.placeholder('title', 'Look up.');
	const action = slide.placeholder(
		'body',
		['Pre-order today at aerie.nimbuswright.example.', 'On roofs from 3 March 2026.'],
		1,
	);
	slide
		.animate(action, {
			entrance: 'fadeIn',
			trigger: 'afterPrevious',
			durationMs: 900,
			delayMs: 400,
		})
		.transition({ type: 'fade', duration: 1200 })
		.notes(
			'Back to the roof at dusk. The weather has always been overhead; from 3 March it is also yours. Thank you. Leave this slide up through the questions.',
		);
}

export function launchSlides(deck: Deck): void {
	deck.section('Launch');
	pricing(deck);
	preOrders(deck);
	close(deck);
}

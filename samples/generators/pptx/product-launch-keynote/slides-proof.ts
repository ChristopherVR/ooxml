/**
 * Slides 7 and 10: the forecast error chart and the platform numbers, with the app and the
 * storm (slides-live.ts) between them.
 */
import type { ChartFontInput } from '../../../../src/core/pptx/core/builders/sdk/index';
import { FONTS } from '../../lib/fiction';
import type { Deck } from '../kit/deck';
import { card, head, tx } from './parts';
import { liveSlides } from './slides-live';
import { C, G, T, body, col, display, label } from './style';
import { LAYOUT } from './template';

const chartFont = (size: number, color: string = C.muted): ChartFontInput => ({
	family: FONTS.body,
	size,
	color,
});

function forecastError(deck: Deck): void {
	const slide = deck.slide(LAYOUT.titleOnly);
	head(slide, 'Proof', 'About half the error, at every lead time');
	const chart = slide.chart(
		'Forecast error by lead time',
		'bar',
		{
			categories: ['6 h', '12 h', '24 h', '36 h', '48 h'],
			series: [
				{ name: 'Regional forecast', values: [1.4, 1.8, 2.3, 2.7, 3.1], color: '#46557A' },
				{
					name: 'Aerie',
					values: [0.6, 0.8, 1.1, 1.4, 1.6],
					color: C.amber,
					dataLabels: {
						showValue: true,
						position: 'outEnd',
						numberFormat: '0.0',
						font: { ...chartFont(10, C.amber), bold: true },
					},
				},
			],
			embedWorkbook: true,
			gapWidth: 70,
			overlap: -8,
			hasLegend: true,
			legendPosition: 't',
			legendFont: chartFont(10.5, C.soft),
			axes: {
				catAx: {
					font: chartFont(10.5),
					line: { color: C.line, width: 1 },
					majorTickMark: 'none',
					minorTickMark: 'none',
					title: 'Hours ahead',
				},
				valAx: {
					font: chartFont(10.5),
					line: 'none',
					majorTickMark: 'none',
					minorTickMark: 'none',
					majorGridlines: true,
					majorGridlinesLine: { color: '#1F2A44', width: 1 },
					numberFormat: '0.0',
					min: 0,
					max: 3.5,
					majorUnit: 0.5,
					title: 'Mean absolute error, °C',
				},
			},
		},
		{ x: G.left - 8, y: 196, width: 728, height: 428 },
	);
	const x = 856;
	const kicker = tx(
		slide,
		'Stat label',
		'AT 24 HOURS',
		{ x, y: 236, width: 344, height: 18 },
		label(),
	);
	const stat = tx(
		slide,
		'Stat',
		'−52 %',
		{ x, y: 262, width: 344, height: 84 },
		display(60, C.amber),
	);
	const reading = tx(
		slide,
		'Stat reading',
		'Aerie misses by 1.1 °C where the regional forecast misses by 2.3 °C.',
		{ x, y: 362, width: 320, height: 80 },
		body(T.body),
		{ lineSpacing: 1.3 },
	);
	tx(
		slide,
		'Source',
		'Mean absolute temperature error at 1,240 beta stations, March to November 2025, against the public 3 km regional model for the same addresses.',
		{ x, y: 540, width: 330, height: 72 },
		body(T.caption, C.muted),
		{ lineSpacing: 1.3 },
	);
	slide
		.animate(chart, {
			entrance: 'wipeIn',
			direction: 'fromBottom',
			trigger: 'afterPrevious',
			durationMs: 900,
		})
		.animate(stat, { entrance: 'zoomIn', trigger: 'onClick', durationMs: 500 })
		.animate(kicker, { entrance: 'fadeIn', trigger: 'withPrevious', durationMs: 500 })
		.animate(reading, { entrance: 'fadeIn', trigger: 'afterPrevious', durationMs: 500 })
		.transition({ type: 'split', direction: 'out', duration: 700 })
		.notes(
			'This is the slide the meteorologists asked for. Grey is the public regional forecast, amber is Aerie, for the same 1,240 addresses over nine months. At every lead time the error is roughly halved; at 24 hours it is 1.1 degrees against 2.3, which is 52 percent less. Click once to bring up the number.',
		);
}

const NUMBERS = [
	['Beta stations', '1,240', 'Roofs in five regions, reporting since March 2025.'],
	['Readings a day', '42.8 M', 'Each station reads its sensors 34,560 times a day.'],
	['Median warning', '22 min', 'Between the alert and the first gust at the door.'],
] as const;

function platform(deck: Deck): void {
	const slide = deck.slide(LAYOUT.titleOnly);
	head(slide, 'The platform', 'Every roof makes the next forecast better');
	const stats = NUMBERS.map(([name, value, text], index) => {
		const { x, width } = col(index * 4, 4);
		const top = 216;
		const parts = [
			card(slide, `Number ${index + 1}`, { x, y: top, width, height: 252 }),
			tx(
				slide,
				`Number ${index + 1} label`,
				name.toUpperCase(),
				{ x: x + 28, y: top + 34, width: width - 56, height: 18 },
				label(index === 1 ? C.amber : C.sky),
			),
			tx(
				slide,
				`Number ${index + 1} text`,
				text,
				{ x: x + 28, y: top + 162, width: width - 56, height: 60 },
				body(T.small + 1),
				{ lineSpacing: 1.3 },
			),
		];
		const value48 = tx(
			slide,
			`Number ${index + 1} value`,
			value,
			{ x: x + 28, y: top + 70, width: width - 56, height: 72 },
			display(T.stat),
		);
		for (const [at, element] of [...parts, value48].entries())
			slide.animate(element, {
				entrance: 'fadeIn',
				trigger: index === 0 && at === 0 ? 'afterPrevious' : 'withPrevious',
				durationMs: 500,
				delayMs: index * 300,
			});
		return value48;
	});
	tx(
		slide,
		'Pooling',
		'Readings are pooled, anonymised and fed back into every forecast within ten minutes.',
		{ x: G.left, y: 512, width: 760, height: 60 },
		body(T.lead - 2, C.muted),
		{ lineSpacing: 1.3 },
	);
	stats.forEach((stat, index) =>
		slide.animate(stat, {
			emphasis: 'growShrink',
			trigger: index === 0 ? 'afterPrevious' : 'withPrevious',
			durationMs: 600,
			delayMs: index * 200,
		}),
	);
	slide
		.transition({ type: 'zoom', direction: 'in', duration: 700 })
		.notes(
			'The station is the product people hold, but the platform is why it gets better. 1,240 beta stations, each reading 34,560 times a day, is 42.8 million readings a day. The three cards arrive one after another and the numbers swell once: with-previous and after-previous timing, no clicks needed.',
		);
}

export function proofSlides(deck: Deck): void {
	deck.section('Proof');
	forecastError(deck);
	liveSlides(deck);
	platform(deck);
}

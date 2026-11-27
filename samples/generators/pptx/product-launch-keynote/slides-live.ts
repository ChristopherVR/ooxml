/**
 * Slides 8 and 9: the app (text built paragraph by paragraph) and the storm (a motion path and
 * a click-triggered alert).
 */
import type { Deck } from '../kit/deck';
import { dataUrl } from '../kit/slide';
import { APP_SIZE, APP_ZONES, appScreen } from './art-product';
import { pressureMap } from './art-scenes';
import { head, paragraphs, tx } from './parts';
import { C, G, T, body, cardFill, display, label, noLine, softShadow } from './style';
import { LAYOUT } from './template';

const APP_POINTS = [
	'A 48-hour forecast for your address, refreshed every ten minutes',
	'Rain start and stop times, to the minute',
	'Frost, wind and storm alerts before the weather arrives',
	'Share your station with the neighbours, or keep it private',
];

function app(deck: Deck): void {
	const slide = deck.slide(LAYOUT.split);
	head(slide, 'The app', 'Your street, on your phone');
	slide.placeholder('body', APP_POINTS, 1);
	const k = 0.74;
	const x = 904;
	const y = 84;
	const at = (px: number, py: number) => ({ x: x + px * k, y: y + py * k });
	const width = APP_SIZE.width * k;
	slide.shape('Phone', 'roundRect', {
		x: x - 10,
		y: y - 10,
		width: width + 20,
		height: APP_SIZE.height * k + 20,
		fill: { type: 'solid', color: '#04060C' },
		stroke: { color: '#2F3B5C', width: 1.5 },
		shadow: softShadow,
		adjustments: { adj: 12500 },
	});
	slide.shape('App screen', 'roundRect', {
		x,
		y,
		width,
		height: APP_SIZE.height * k,
		fill: { type: 'image', url: dataUrl(appScreen()) },
		stroke: noLine,
		adjustments: { adj: 9800 },
	});
	const line = (
		name: string,
		text: string,
		px: number,
		py: number,
		style: ReturnType<typeof body>,
	) => tx(slide, name, text, { ...at(px, py), width: width - 2 * px * k, height: 20 }, style);
	line('App place', 'LARKSPUR HEIGHTS', 24, 40, { ...label(C.soft), fontSize: 8.6 });
	line('App temperature', '7.4°', 24, 66, display(34));
	line('App feels', 'Feels like 5.1°. Wind 18 km/h, north-west.', 24, 164, body(7.5));
	line('App chart label', 'NEXT 48 HOURS', 24, 236, { ...label(), fontSize: 8 });
	line('App rain label', 'RAIN STARTS 16:42', 24, 416, { ...label(C.sky), fontSize: 8 });
	const cards = [
		['Frost tonight', 'Cover the seedlings by 22:00'],
		['Squall line, 22 minutes away', 'Gusts to 74 km/h from the north-west'],
	] as const;
	cards.forEach(([heading, text], index) => {
		const zone = APP_ZONES.cards[index]!;
		slide.shape(`App card ${index + 1}`, 'rect', {
			...at(78, zone.top),
			width: 250 * k,
			height: (zone.bottom - zone.top) * k,
			stroke: noLine,
			insets: { left: 0, top: 0, right: 0, bottom: 0 },
			textStyle: { alignment: 'left', verticalAlignment: 'middle' },
			text: paragraphs([{ text: heading, style: display(8) }], [{ text, style: body(7) }]),
		});
	});
	const squall = APP_ZONES.cards[1];
	slide.shape('Alert highlight', 'roundRect', {
		...at(20, squall.top),
		width: (APP_SIZE.width - 40) * k,
		height: (squall.bottom - squall.top) * k,
		stroke: { color: C.amber, width: 2 },
		glow: { color: C.amber, radius: 10, opacity: 0.45 },
		adjustments: { adj: 18000 },
	});
	// Placeholders are named by type and order: kicker, title, number, then this body.
	const list = 'Body 4';
	APP_POINTS.forEach((_, paragraph) =>
		slide.effects.push({ target: list, paragraph, kind: 'rise', duration: 500 }),
	);
	// The alert in the phone lights up together with the line about alerts.
	slide.effects.splice(3, 0, {
		target: 'Alert highlight',
		kind: 'fade',
		start: 'with',
		duration: 700,
	});
	slide
		.transition({ type: 'push', direction: 'u', duration: 700 })
		.notes(
			'Four clicks, one line each. The home screen is deliberately quiet: the temperature on your roof, the next 48 hours, and when the rain starts. On the third click the squall card in the phone lights up with the line about alerts. Everything on the phone is live text over a generated picture, so it can be localised.',
		);
}

function storm(deck: Deck): void {
	const slide = deck.slide(LAYOUT.titleOnly);
	head(slide, 'Storm mode', 'It saw the squall 22 minutes out');
	const map = { x: G.left, y: 204, width: 840, height: 420 };
	slide.picture('Pressure map', pressureMap(), {
		...map,
		altText: 'A pressure map with isobars closing around a low to the west of the coast',
	});
	const station = { x: map.x + map.width * 0.7, y: map.y + map.height * 0.52 };
	const front = slide.freeform(
		'Squall line',
		'M 30 0 C 130 90 130 250 30 340 C 80 250 80 90 30 0 Z',
		{
			x: 150,
			y: 244,
			width: 130,
			height: 340,
			fill: {
				type: 'gradient',
				angle: 0,
				stops: [
					{ color: C.violet, position: 0, opacity: 0.15 },
					{ color: C.sky, position: 100, opacity: 0.95 },
				],
			},
			stroke: noLine,
			glow: { color: C.sky, radius: 14, opacity: 0.35 },
		},
	);
	const ring = slide.shape('Station ring', 'ellipse', {
		x: station.x - 20,
		y: station.y - 20,
		width: 40,
		height: 40,
		stroke: { color: C.amber, width: 2 },
	});
	slide.shape('Station', 'ellipse', {
		x: station.x - 7,
		y: station.y - 7,
		width: 14,
		height: 14,
		fill: { type: 'solid', color: C.amber },
		stroke: noLine,
		glow: { color: C.amber, radius: 12, opacity: 0.5 },
	});
	tx(
		slide,
		'Station label',
		'YOUR ROOF',
		{ x: station.x + 30, y: station.y - 9, width: 140, height: 18 },
		label(C.text),
	);
	const x = 952;
	const alert = slide.shape('Alert card', 'roundRect', {
		x,
		y: 204,
		width: 248,
		height: 184,
		fill: cardFill,
		stroke: { color: C.coral, width: 1.5 },
		shadow: softShadow,
		adjustments: { adj: 7000 },
		insets: { left: 22, top: 0, right: 18, bottom: 0 },
		textStyle: { alignment: 'left', verticalAlignment: 'middle', spaceAfter: 6 },
		text: paragraphs(
			[{ text: 'SQUALL WARNING', style: label(C.coral) }],
			[{ text: 'Gusts to 74 km/h', style: display(17) }],
			[{ text: 'Reaches your roof at 16:42, in 22 minutes.', style: body(T.small) }],
		),
	});
	tx(
		slide,
		'Log',
		[
			[
				{ text: '16:20  ', style: { color: C.amber, bold: true } },
				{ text: 'Pressure falls 1.8 hPa in six minutes' },
			],
			[
				{ text: '16:42  ', style: { color: C.amber, bold: true } },
				{ text: 'The squall line crosses Larkspur Heights' },
			],
		],
		{ x, y: 416, width: 248, height: 110 },
		body(T.small),
		{ lineSpacing: 1.3, spaceAfter: 10 },
	);
	tx(
		slide,
		'Hint',
		'CLICK TO REHEARSE THE ALERT',
		{ x, y: 544, width: 248, height: 16 },
		{
			...label(),
			fontSize: 10.6,
		},
	);
	const pill = slide.shape('Storm mode', 'roundRect', {
		x,
		y: 572,
		width: 248,
		height: 52,
		fill: {
			type: 'gradient',
			angle: 0,
			stops: [
				{ color: C.amber, position: 0 },
				{ color: C.coral, position: 100 },
			],
		},
		stroke: noLine,
		shadow: softShadow,
		adjustments: { adj: 50000 },
		textStyle: { alignment: 'center', verticalAlignment: 'middle' },
		text: [{ text: 'Storm mode', style: display(14, C.night) }],
	});
	slide
		.animate(front, { entrance: 'fadeIn', trigger: 'afterPrevious', durationMs: 600 })
		.animate(front, {
			motionPath: 'M 0 0 C 0.09 -0.02 0.2 0.03 0.275 0.01 E',
			trigger: 'afterPrevious',
			durationMs: 3200,
		})
		.animate(ring, { emphasis: 'pulse', trigger: 'afterPrevious', durationMs: 700, repeatCount: 3 })
		.animate(alert, {
			entrance: 'zoomIn',
			trigger: 'onShapeClick',
			triggerShapeId: pill.id,
			durationMs: 450,
		})
		.transition({ type: 'cover', direction: 'l', duration: 700 })
		.notes(
			'A real afternoon from the beta, 9 October in Larkspur Heights. The squall line travels across the pressure map on a motion path and the station ring pulses when it arrives. Click the Storm mode button to show the warning the owner received at 16:20, 22 minutes before the gusts. The warning is a triggered animation, so it can be shown at any point.',
		);
}

export function liveSlides(deck: Deck): void {
	app(deck);
	storm(deck);
}

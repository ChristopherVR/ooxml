/**
 * Slides 1 to 4: the cover, the problem, the reveal and the three promises. Three shapes carry
 * the Morph transitions between them: `!!glow` (the last light, then the light of the ring),
 * `!!accent` (a rule that becomes a distance line and a card edge) and `!!device`.
 */
import type { Deck } from '../kit/deck';
import { pt } from '../kit/geometry';
import { DEVICE_SIZE, deviceRender } from './art-product';
import { card, head, paragraphs, restyle, tx } from './parts';
import { C, T, body, display, glowFill, label, noLine } from './style';
import { LAYOUT } from './template';

const DEVICE_RATIO = DEVICE_SIZE.width / DEVICE_SIZE.height;
/** The render is drawn once and used on two slides. */
let rendered: Uint8Array | undefined;
const device = () => (rendered ??= deviceRender());
/** Where the lit ring sits in the device picture, as fractions of its box. */
const RING = { u: 0.5, v: 276 / DEVICE_SIZE.height };

function cover(deck: Deck): void {
	const slide = deck.slide(LAYOUT.cover);
	slide.shape('!!glow', 'ellipse', {
		x: 767,
		y: 295,
		width: 360,
		height: 360,
		fill: glowFill(C.amber, 0.3),
		stroke: noLine,
	});
	slide.placeholder('ctrTitle', ['The weather,', 'from your', 'own roof.']);
	const subtitle = slide.placeholder(
		'subTitle',
		'Introducing Aerie, the home weather station from Nimbuswright. Port Averly, 15 January 2026.',
	);
	const accent = slide.shape('!!accent', 'rect', {
		x: 80,
		y: 612,
		width: 64,
		height: 5,
		fill: { type: 'solid', color: C.amber },
		stroke: noLine,
	});
	slide
		.animate(subtitle, {
			entrance: 'fadeIn',
			trigger: 'afterPrevious',
			durationMs: 900,
			delayMs: 300,
		})
		.animate(accent, {
			entrance: 'wipeIn',
			direction: 'fromLeft',
			trigger: 'afterPrevious',
			durationMs: 500,
		})
		.transition({ type: 'fade', duration: 900 })
		.notes(
			'Good evening, and thank you for coming up to the roof with us. For four years Nimbuswright has built forecasting software for people who could not afford to be wrong about the weather. Tonight we put that in a box you can carry up a ladder. Let the cover sit for a moment before the first click.',
		);
}

function problem(deck: Deck): void {
	const slide = deck.slide(LAYOUT.statement);
	slide.placeholder('body', 'The problem', 10);
	const statement = slide.placeholder('title', [
		'Your forecast was made',
		'for an airport',
		'14 km away.',
	]);
	restyle(statement, '14 km away.', { color: C.amber, colorRef: { scheme: 'accent1' } });
	slide.placeholder(
		'body',
		'A regional model paints a whole city with one brush. Your street has its own wind, its own rain and its own frost.',
		11,
	);
	slide.placeholder('sldNum', '2', 12);
	// The distance, drawn: the forecast point on the left, your roof on the right.
	const y = 566;
	slide.shape('!!glow', 'ellipse', {
		x: 1100,
		y: y - 70,
		width: 140,
		height: 140,
		fill: glowFill(C.amber, 0.5),
		stroke: noLine,
	});
	slide.shape('!!accent', 'rect', {
		x: 884,
		y: y - 1,
		width: 262,
		height: 2,
		fill: { type: 'solid', color: C.amber },
		stroke: noLine,
	});
	slide.shape('Forecast point', 'ellipse', {
		x: 860,
		y: y - 8,
		width: 16,
		height: 16,
		stroke: { color: C.muted, width: 2 },
	});
	slide.freeform('Your roof', 'M 0 26 L 0 12 L 15 0 L 30 12 L 30 26 Z', {
		x: 1155,
		y: y - 15,
		width: 30,
		height: 26,
		fill: { type: 'solid', color: C.amber },
		stroke: noLine,
	});
	tx(
		slide,
		'Distance',
		'14 km',
		{ x: 884, y: y - 34, width: 262, height: 22 },
		display(13, C.amber),
		{
			alignment: 'center',
		},
	);
	tx(
		slide,
		'Forecast label',
		'KESSINGTON AIRFIELD',
		{ x: 800, y: y + 22, width: 200, height: 18 },
		label(),
	);
	tx(
		slide,
		'Roof label',
		'YOUR ROOF',
		{ x: 1000, y: y + 22, width: 200, height: 18 },
		label(C.soft),
		{
			alignment: 'right',
		},
	);
	slide
		.transition({ type: 'morph', duration: 1400 })
		.notes(
			'Every forecast you read this morning was computed for a grid point, and for most of Port Averly that point is the airfield at Kessington, 14 km inland. It is a good forecast for the airfield. The amber line on the right is the gap we are here to close. The glow and the rule from the cover travel into this drawing: that is the Morph transition.',
		);
}

function reveal(deck: Deck): void {
	const slide = deck.slide(LAYOUT.hero);
	const height = 600;
	const width = height * DEVICE_RATIO;
	const x = 742;
	const y = 34;
	slide.shape('!!glow', 'ellipse', {
		x: x + width * RING.u - 350,
		y: y + height * RING.v - 350,
		width: 700,
		height: 700,
		fill: glowFill(C.amber, 0.42),
		stroke: noLine,
	});
	slide.picture('!!device', device(), {
		x,
		y,
		width,
		height,
		altText:
			'The Aerie weather station: a louvred pod with a lit amber ring, wind cups and a solar panel',
	});
	slide.placeholder('body', 'The product', 10);
	slide.placeholder('title', 'Meet Aerie.');
	const lead = slide.placeholder(
		'body',
		'A complete weather station the size of a coffee tin. On your roof in ten minutes, on your phone in eleven.',
		11,
	);
	slide.placeholder('sldNum', '3', 12);
	slide.shape('!!accent', 'rect', {
		x: 80,
		y: 452,
		width: 64,
		height: 5,
		fill: { type: 'solid', color: C.amber },
		stroke: noLine,
	});
	slide
		.animate(lead, { entrance: 'fadeIn', trigger: 'afterPrevious', durationMs: 800 })
		.transition({ type: 'morph', duration: 1600 })
		.notes(
			'This is Aerie. One pod, one mast, no wires: the ring you see lit is the sensor band, and the small panel keeps it running through three weeks without sun. Pause here. The glow from the little roof on the last slide grows into the light of the ring.',
		);
}

const PROMISES = [
	[
		'It measures your roof, not the airfield',
		'Eleven signals, sampled every 2.5 seconds, exactly where you live.',
	],
	[
		'It learns your microclimate',
		'Fourteen days of readings teach the model how your street bends the regional forecast.',
	],
	[
		'It warns you first',
		'Squalls, frost and downpours: a median 22 minutes before they reach your door.',
	],
] as const;

function promises(deck: Deck): void {
	const slide = deck.slide(LAYOUT.titleOnly);
	head(slide, 'What it does', 'Three things a regional forecast cannot do');
	const height = 430;
	const width = height * DEVICE_RATIO;
	const x = 96;
	const y = 204;
	slide.shape('!!glow', 'ellipse', {
		x: x + width * RING.u - 230,
		y: y + height * RING.v - 230,
		width: 460,
		height: 460,
		fill: glowFill(C.sky, 0.3),
		stroke: noLine,
	});
	slide.picture('!!device', device(), {
		x,
		y,
		width,
		height,
		altText: 'The Aerie weather station',
	});
	const left = 468;
	PROMISES.forEach(([heading, text], index) => {
		const top = 204 + index * 144;
		const panel = card(slide, `Promise ${index + 1}`, { x: left, y: top, width: 732, height: 124 });
		const copy = slide.shape(`Promise ${index + 1} text`, 'rect', {
			x: left + 112,
			y: top,
			width: 592,
			height: 124,
			stroke: noLine,
			insets: { left: 0, top: 0, right: 0, bottom: 0 },
			textStyle: { alignment: 'left', verticalAlignment: 'middle', spaceAfter: 4 },
			text: paragraphs([{ text: heading, style: display(T.lead) }], [{ text, style: body() }]),
		});
		const numeral = tx(
			slide,
			`Promise ${index + 1} number`,
			`0${index + 1}`,
			{ x: left + 32, y: top, width: 70, height: 124 },
			{ ...display(26, C.amber), fontSize: pt(26) },
			{ verticalAlignment: 'middle' },
		);
		for (const [at, element] of [panel, copy, numeral].entries())
			slide.animate(element, {
				entrance: 'fadeIn',
				trigger: index === 0 && at === 0 ? 'afterPrevious' : 'withPrevious',
				durationMs: 500,
				delayMs: index * 280,
			});
	});
	slide.shape('!!accent', 'rect', {
		x: left,
		y: 224,
		width: 4,
		height: 84,
		fill: { type: 'solid', color: C.amber },
		stroke: noLine,
	});
	slide
		.transition({ type: 'morph', duration: 1200 })
		.notes(
			'Aerie steps aside and its light cools to the blue of a working instrument. Three promises, and each one is a number we will defend in the next ten minutes: eleven signals every 2.5 seconds, fourteen days to learn a street, and a median warning of 22 minutes.',
		);
}

export function openingSlides(deck: Deck): void {
	deck.section('Opening');
	cover(deck);
	problem(deck);
	deck.section('The product');
	reveal(deck);
	promises(deck);
}

/** Section "Why", first half: the extension on the network map, and the three decisions. */
import type { PptxElement } from '../../../../src/core/pptx/core/types/elements';
import type { Deck } from '../kit/deck';
import { SLIDE_H, SLIDE_W } from '../kit/geometry';
import type { Slide } from '../kit/slide';
import { MAP_H, MAP_W, T1, T1_EXTENSION, networkMap, onSlide } from './art';
import {
	FARES,
	FREQUENCY_RATIO,
	TOP_ROUTES,
	TRAM_AFTER,
	TRAM_BEFORE,
	change,
	signed,
} from './data';
import { card, stat } from './parts';
import { C, label, span } from './style';

const MAP = { ...span(0, 8), y: 196 };

/** A PowerPoint motion path through the given slide points, relative to the first one. */
function motionPath(points: { x: number; y: number }[]): string {
	const [first, ...rest] = points;
	const steps = rest.map(
		(point) =>
			`L ${((point.x - first!.x) / SLIDE_W).toFixed(4)} ${((point.y - first!.y) / SLIDE_H).toFixed(4)}`,
	);
	return `M 0 0 ${steps.join(' ')} E`;
}

function mapSlide(deck: Deck): void {
	const slide = deck.slide('Title only');
	slide.placeholder('title', 'Four new stops put the north shore of Harbourside on the tram map');
	slide.placeholder('body', 'Why: the extension', 10);
	slide.placeholder('sldNum', '7', 12);
	slide.picture('Network map', networkMap(), {
		...MAP,
		height: (MAP.width * MAP_H) / MAP_W,
		altText:
			'A map of Port Averly: two tram lines meeting at Old Quay, bus routes, two ferry routes across the harbour, and four new tram stops on the north shore.',
	});
	const place = (
		name: string,
		text: string,
		at: readonly [number, number],
		color: string,
		bold = false,
	) => {
		const point = onSlide(at, MAP);
		return label(
			slide,
			name,
			text,
			{ x: point.x, y: point.y, width: 220, height: 16 },
			{
				size: bold ? 9.5 : 8,
				color,
				bold,
				display: !bold,
				tracking: bold ? 0 : 1.4,
			},
		);
	};
	place('District Northgate', 'NORTHGATE', [330, 60], C.slate);
	place('District Marrow Hill', 'MARROW HILL', [60, 318], C.slate);
	place('District Old Quay', 'OLD QUAY', [372, 296], C.slate);
	place('District Harbourside', 'HARBOURSIDE', [640, 236], C.slate);
	place('District Saltmeadow', 'SALTMEADOW', [250, 548], C.slate);
	place('Line T1', 'T1 Harbour Line', [96, 448], C.coral, true);
	place('Line T2', 'T2 Northgate Line', [150, 66], C.teal, true);
	place('Line F1', 'F1 Harbour Ferry', [742, 446], '#A87A1F', true);
	const newStops = label(
		slide,
		'New stops',
		'Four new stops, opened 1 September 2025',
		{ ...onSlide([590, 34], MAP), width: 190, height: 34 },
		{ size: 10, color: C.coral, bold: true, align: 'right', leading: 1.25 },
	);
	const start = onSlide(T1[0]!, MAP);
	const tram = slide.shape('Tram', 'ellipse', {
		x: start.x - 9,
		y: start.y - 9,
		width: 18,
		height: 18,
		fill: { type: 'solid', color: C.ink },
		stroke: { color: C.white, width: 2.5 },
		shadow: { color: '#000000', blur: 8, offsetX: 0, offsetY: 3, opacity: 0.3 },
	});
	const growth = signed(change(TOP_ROUTES[0]!.boardings, TOP_ROUTES[0]!.before));
	const side = span(9, 3);
	stat(
		slide,
		'Stat T1',
		growth,
		'boardings on T1 Harbour Line, the largest gain of any route',
		{
			...side,
			y: 196,
			height: 130,
		},
		C.coral,
	);
	stat(slide, 'Stat stops', '4', 'new stops and 2.3 km of track along the north shore', {
		...side,
		y: 344,
		height: 130,
	});
	stat(slide, 'Stat reach', '31,000', 'more residents within a ten-minute walk of a tram stop', {
		...side,
		y: 492,
		height: 130,
	});
	const route = [...T1, ...T1_EXTENSION.slice(1)].map((point) => onSlide(point, MAP));
	slide.animate(tram, {
		motionPath: motionPath(route),
		trigger: 'afterPrevious',
		durationMs: 5000,
	});
	slide.animate(newStops, { entrance: 'fadeIn', trigger: 'afterPrevious', durationMs: 500 });
	slide.animate(newStops, { emphasis: 'pulse', trigger: 'afterPrevious', durationMs: 700 });
	slide.transition({ type: 'cover', direction: 'l', duration: 700 });
	slide.notes(
		`This is the network. The coral line is T1, the Harbour Line; watch the tram run its length and on to the four stops that opened on 1 September. The extension added 2.3 km of track and put about 31,000 more residents within a ten-minute walk of a tram stop. T1 boardings grew ${growth} over the year.`,
	);
}

interface Decision {
	date: string;
	title: string;
	body: string;
	figure: string;
	caption: string;
	color: string;
	/** A glyph drawn in a 24 x 24 box. */
	glyph: string;
}

function decisionCard(slide: Slide, decision: Decision, column: number): PptxElement[] {
	const box = { ...span(column * 4, 4), y: 204, height: 404 };
	const x = box.x + 28;
	const width = box.width - 56;
	return [
		card(slide, `${decision.title} card`, box),
		slide.shape(`${decision.title} badge`, 'ellipse', {
			x: box.x + box.width - 28 - 48,
			y: box.y + 28,
			width: 48,
			height: 48,
			fill: {
				type: 'gradient',
				angle: 45,
				stops: [
					{ color: decision.color, position: 0 },
					{ color: decision.color === C.coral ? C.ochre : C.tealSoft, position: 100 },
				],
			},
			stroke: { color: 'transparent' },
		}),
		slide.freeform(`${decision.title} glyph`, decision.glyph, {
			x: box.x + box.width - 28 - 36,
			y: box.y + 40,
			width: 24,
			height: 24,
			fill: { type: 'none' },
			stroke: { color: C.white, width: 2.25, cap: 'rnd', join: 'round' },
		}),
		label(
			slide,
			`${decision.title} date`,
			decision.date,
			{ x, y: box.y + 32, width: width - 60, height: 16 },
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
			`${decision.title} title`,
			decision.title,
			{ x, y: box.y + 56, width: width - 60, height: 60 },
			{
				size: 20,
				display: true,
				bold: true,
				leading: 1.1,
			},
		),
		label(
			slide,
			`${decision.title} body`,
			decision.body,
			{ x, y: box.y + 132, width, height: 96 },
			{
				size: 12,
				color: C.slate,
				leading: 1.4,
			},
		),
		slide.shape(`${decision.title} rule`, 'rect', {
			x,
			y: box.y + 252,
			width,
			height: 1,
			fill: { type: 'solid', color: C.rule },
			stroke: { color: 'transparent' },
		}),
		label(
			slide,
			`${decision.title} figure`,
			decision.figure,
			{ x, y: box.y + 272, width, height: 58 },
			{
				size: 44,
				display: true,
				bold: true,
				color: decision.color,
				tracking: -1,
				leading: 1,
			},
		),
		label(
			slide,
			`${decision.title} caption`,
			decision.caption,
			{ x, y: box.y + 338, width, height: 40 },
			{
				size: 11,
				color: C.slate,
				leading: 1.3,
			},
		),
	];
}

function decisionsSlide(deck: Deck): void {
	const slide = deck.slide('Title only');
	slide.placeholder('title', 'Three decisions explain most of the change');
	slide.placeholder('body', 'Why', 10);
	slide.placeholder('sldNum', '8', 12);
	const tramGain = signed(change(TRAM_AFTER, TRAM_BEFORE));
	const decisions: Decision[] = [
		{
			date: '3 MARCH',
			title: 'Capped tap-and-go fares',
			body: 'Tap a bank card or phone on any vehicle. Daily and weekly caps apply by themselves, so nobody has to choose a ticket.',
			figure: `${FARES[0]!.share}%`,
			caption: 'of 2025 boardings, the most used fare in its first year',
			color: C.teal,
			glyph:
				'M 5 8 C 7 10.5 7 13.5 5 16 M 10 5.5 C 13 9.5 13 14.5 10 18.5 M 15 3 C 19.5 8.5 19.5 15.5 15 21',
		},
		{
			date: '2 JUNE',
			title: 'The Frequent Network',
			body: 'Six bus routes now run every 8 minutes or better, from 6 am to 9 pm, every day of the week. No timetable needed.',
			figure: `${FREQUENCY_RATIO.toFixed(1)}x`,
			caption: 'the boardings per service hour of our other bus routes',
			color: C.teal,
			glyph: 'M 4 6 L 10 12 L 4 18 M 12 6 L 18 12 L 12 18',
		},
		{
			date: '1 SEPTEMBER',
			title: 'Harbour Line extension',
			body: 'Four new tram stops and 2.3 km of track along the north shore of Harbourside, the fastest growing district in the city.',
			figure: tramGain,
			caption: 'weekly tram boardings, after against before opening',
			color: C.coral,
			glyph: 'M 3 19 L 11 11 L 21 4 M 13 4 L 21 4 L 21 12',
		},
	];
	decisions.forEach((decision, column) => {
		const parts = decisionCard(slide, decision, column);
		parts.forEach((part, index) =>
			slide.animate(part, {
				entrance: 'fadeIn',
				trigger: index > 0 ? 'withPrevious' : column === 0 ? 'afterPrevious' : 'onClick',
				durationMs: 500,
			}),
		);
	});
	slide.transition({ type: 'fade', duration: 600 });
	slide.notes(
		`Three decisions, in the order they took effect. Capped fares in March removed the question of which ticket to buy. The Frequent Network in June made six bus routes turn-up-and-go. The Harbour Line extension in September lifted weekly tram boardings by ${tramGain}. The next three slides take them one at a time. Each click brings in the next card.`,
	);
}

export function whySlides(deck: Deck): void {
	deck.section('Why');
	mapSlide(deck);
	decisionsSlide(deck);
}

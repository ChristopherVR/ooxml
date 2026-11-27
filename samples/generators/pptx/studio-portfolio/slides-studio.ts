/** The opening section: the cover, who the studio is, and the index of the three projects. */
import type { PptxElement } from '../../../../src/core/pptx/core/types/elements';
import { FONTS } from '../../lib/fiction';
import type { Deck } from '../kit/deck';
import { pt } from '../kit/geometry';
import type { Slide } from '../kit/slide';
import { CASES, MIX, PROJECTS, STUDIO } from './content';
import { C, G, T, col, figure, headline, kicker, label, numbered, rule, span } from './style';
import { ALT, png, type PictureName } from './pictures';
import { LAYOUT } from './template';

/** The hero picture of each case study, in case order. */
export const HEROES: PictureName[] = ['hero1', 'hero2', 'hero3'];

/** The studio's mark as SVG path data in a `width` x `height` box: a thimble in silhouette. */
export function thimblePath(width: number, height: number): string {
	const x = (value: number) => ((value * width) / 100).toFixed(2);
	const y = (value: number) => ((value * height) / 100).toFixed(2);
	return (
		`M ${x(8)} ${y(100)} L ${x(92)} ${y(100)} L ${x(92)} ${y(86)} L ${x(84)} ${y(86)} ` +
		`L ${x(74)} ${y(30)} C ${x(72)} ${y(-8)} ${x(28)} ${y(-8)} ${x(26)} ${y(30)} ` +
		`L ${x(16)} ${y(86)} L ${x(8)} ${y(86)} Z`
	);
}

/** The mark and the studio's name beside it. */
export function mark(slide: Slide, x: number, y: number, ink: string): PptxElement {
	const thimble = slide.freeform('Studio mark', thimblePath(26, 30), {
		x,
		y,
		width: 26,
		height: 30,
		fill: { type: 'solid', color: C.vermilion },
		stroke: { color: 'transparent' },
	});
	label(
		slide,
		'Studio name',
		[[{ text: 'Thimblewick Studio', style: { characterSpacing: -0.2 } }]],
		{
			x: x + 38,
			y: y + 6,
			width: 260,
			height: 22,
			fontFamily: FONTS.display,
			fontSize: pt(13),
			bold: true,
			color: ink,
		},
	);
	return thimble;
}

function cover(deck: Deck): void {
	const slide = deck.slide(LAYOUT.cover);
	const thimble = mark(slide, G.left, G.top, C.ink);
	const picture = slide.picture('Cover composition', png('cover'), {
		x: 760,
		y: G.top,
		width: 440,
		height: 506,
		altText: ALT.cover,
	});
	const tag = kicker(slide, 'Cover kicker', 'Portfolio', G.left, 262, 300, C.vermilion);
	slide.placeholder('ctrTitle', ['Selected', 'work']);
	slide.placeholder(
		'subTitle',
		'Identity, wayfinding and editorial design, 2023 to 2025. Three projects in full, and how we work.',
	);
	numbered(slide);
	slide.animate(thimble, { entrance: 'zoomIn', trigger: 'afterPrevious', durationMs: 500 });
	slide.animate(tag, { entrance: 'fadeIn', trigger: 'withPrevious', durationMs: 500 });
	slide.animate(picture, {
		entrance: 'wipeIn',
		direction: 'fromBottom',
		trigger: 'afterPrevious',
		durationMs: 900,
	});
	slide.transition({ type: 'fade', duration: 700 });
	slide.notes(
		`Welcome. This is a selection of what Thimblewick Studio made between 2023 and 2025: three projects shown from brief to result, then a short look at how we work. We are ${STUDIO.people} people in ${STUDIO.city}, so everything here was made by the people you will meet at the end.`,
	);
}

const MIX_COLORS = [C.vermilion, C.blue, C.sage, C.wheat];

function studio(deck: Deck): void {
	const slide = deck.slide(LAYOUT.statement);
	slide.placeholder('title', 'A small studio for things that last.');
	slide.placeholder(
		'body',
		[
			`We are ${STUDIO.people} designers in ${STUDIO.city} who make identities, signs and books.`,
			'We work for people who make real things: bread, gardens, boots.',
			'We stay until the last sign is on the wall and the files are in your hands.',
		],
		1,
	);
	numbered(slide);
	// Facts along the bottom of the text column.
	rule(slide, 'Facts rule', G.left, 500, span(6));
	const facts: [string, string][] = [
		[String(STUDIO.people), 'people, all of them designers or makers'],
		[String(STUDIO.founded), `founded on Ropewalk Lane, ${STUDIO.city}`],
		[String(PROJECTS), 'projects shipped from 2023 to 2025'],
	];
	facts.forEach(([value, caption], index) =>
		figure(slide, `Fact ${index + 1}`, col(index * 2), 520, span(2) - 12, value, caption),
	);
	// The project mix, as a doughnut with its own legend.
	const left = col(7) + 36;
	kicker(slide, 'Mix kicker', 'Projects by discipline', left, 72, 320);
	slide.chart(
		'Project mix chart',
		'doughnut',
		{
			categories: MIX.map(([discipline]) => discipline),
			series: [
				{ name: 'Projects', values: MIX.map(([, count]) => count), pointColors: MIX_COLORS },
			],
			hasLegend: false,
			holeSize: 70,
			embedWorkbook: true,
		},
		{ x: left - 10, y: 104, width: 340, height: 340 },
	);
	headline(slide, 'Mix total', String(PROJECTS), {
		x: left + 60,
		y: 236,
		width: 200,
		height: 60,
		size: 46,
		alignment: 'center',
	});
	label(slide, 'Mix total caption', 'projects', {
		x: left + 60,
		y: 296,
		width: 200,
		height: 20,
		fontSize: pt(T.caption),
		color: C.soft,
		alignment: 'center',
	});
	MIX.forEach(([discipline, count], index) => {
		const x = left + (index % 2) * 170;
		const y = 470 + Math.floor(index / 2) * 34;
		slide.shape(`Legend swatch ${index + 1}`, 'ellipse', {
			x,
			y: y + 4,
			width: 10,
			height: 10,
			fill: { type: 'solid', color: MIX_COLORS[index]! },
			stroke: { color: 'transparent' },
		});
		label(
			slide,
			`Legend ${index + 1}`,
			[[{ text: `${count}  `, style: { bold: true } }, { text: discipline }]],
			{ x: x + 20, y, width: 140, height: 20, fontSize: pt(T.body - 1) },
		);
	});
	// A real paragraph build: one statement per click, then the chart.
	slide.effects.push(
		{ target: 'Body 2', paragraph: 0, kind: 'rise', duration: 600 },
		{ target: 'Body 2', paragraph: 1, kind: 'rise', duration: 600 },
		{ target: 'Body 2', paragraph: 2, kind: 'rise', duration: 600 },
		{ target: 'Project mix chart', kind: 'fade', start: 'after', duration: 700 },
		{ target: 'Mix total', kind: 'fade', start: 'with', duration: 700 },
		{ target: 'Mix total caption', kind: 'fade', start: 'with', duration: 700 },
	);
	slide.transition({ type: 'push', direction: 'l', duration: 700 });
	slide.notes(
		`Three sentences sum us up, and they build one at a time. We are ${STUDIO.people} people, founded in ${STUDIO.founded}. Since 2023 we have shipped ${PROJECTS} projects: ${MIX.map(([discipline, count]) => `${count} ${discipline.toLowerCase()}`).join(', ')}. Identity is the largest share, but wayfinding and editorial are where we have grown.`,
	);
}

function index(deck: Deck): void {
	const slide = deck.slide(LAYOUT.canvas);
	slide.placeholder('title', 'Three projects, from brief to result.');
	numbered(slide);
	const width = span(4) - 16;
	const height = width / 1.6;
	const top = 200;
	const thumbs: PptxElement[] = [];
	const teasers: PptxElement[] = [];
	CASES.forEach((study, at) => {
		const x = col(at * 4) + 8;
		const below = top + height + 30;
		// A paper mount with a soft shadow, so each thumbnail reads as a print on the wall.
		slide.shape(`Mount ${study.number}`, 'rect', {
			x: x - 8,
			y: top - 8,
			width: width + 16,
			height: height + 16,
			fill: { type: 'solid', color: '#FFFFFF' },
			stroke: { color: 'transparent' },
			shadow: { color: '#3A2F22', blur: 24, offsetY: 8, opacity: 0.16 },
		});
		thumbs.push(
			slide.picture(`!!case${at + 1} image`, png(HEROES[at]!), {
				x,
				y: top,
				width,
				height,
				altText: ALT[HEROES[at]!],
			}),
		);
		headline(slide, `!!case${at + 1} number`, study.number, {
			x: x - 8,
			y: below - 3,
			width: 40,
			height: 24,
			size: 15,
			color: C.vermilion,
		});
		kicker(
			slide,
			`Meta ${study.number}`,
			`${study.discipline}, ${study.year}`,
			x + 28,
			below + 2,
			width - 28,
		);
		headline(slide, `Client ${study.number}`, study.client, {
			x: x - 8,
			y: below + 28,
			width: width + 8,
			height: 48,
			size: 15,
			lineSpacing: 1.05,
		});
		teasers.push(
			label(slide, `Teaser ${study.number}`, study.teaser, {
				x: x - 8,
				y: below + 82,
				width: width - 24,
				height: 44,
				fontSize: pt(T.body - 1),
				color: C.soft,
				lineSpacing: 1.3,
			}),
		);
	});
	label(slide, 'Hint', 'Select a picture for the one-line summary.', {
		x: G.left,
		y: 152,
		width: 500,
		height: 18,
		fontSize: pt(T.caption),
		color: C.soft,
	});
	thumbs.forEach((thumb, at) => {
		slide.animate(thumb, {
			entrance: 'wipeIn',
			direction: 'fromLeft',
			trigger: 'afterPrevious',
			durationMs: 500,
		});
		// The interactive sequence: clicking a picture reveals that project's summary.
		slide.animate(teasers[at]!, {
			entrance: 'fadeIn',
			trigger: 'onShapeClick',
			triggerShapeId: thumb.id,
			durationMs: 500,
		});
	});
	slide.transition({ type: 'wipe', direction: 'l', duration: 700 });
	slide.notes(
		`The three projects we will walk through: ${CASES.map((study) => `${study.client} (${study.discipline.toLowerCase()}, ${study.year})`).join('; ')}. In a slide show, click any picture to reveal its one-line summary. The first picture grows into the next slide.`,
	);
}

export function studioSlides(deck: Deck): void {
	deck.section('Studio');
	cover(deck);
	studio(deck);
	deck.section('Selected work');
	index(deck);
}

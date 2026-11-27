/** The closing slides: results across the projects, how the studio works, the team, contact. */
import type {
	TableCellInput,
	TableRowInput,
} from '../../../../src/core/pptx/core/builders/sdk/index';
import { applyBuiltinSmartArtLayout } from '../../../../src/core/pptx/smartart-layouts/index';
import type { PptxSmartArtData } from '../../../../src/core/pptx/core/types/smart-art';
import { FONTS } from '../../lib/fiction';
import type { Deck } from '../kit/deck';
import { pt } from '../kit/geometry';
import { CASES, PHASES, STUDIO, TEAM, TOTAL_WEEKS } from './content';
import { ALT, png, type PictureName } from './pictures';
import { mark } from './slides-studio';
import { C, G, T, col, headline, kicker, label, numbered, rule, span } from './style';
import { LAYOUT } from './template';

const HAIRLINE = { color: C.rule, width: 1 };
const HEAVY = { color: C.ink, width: 1.5 };
/**
 * The SDK cannot write "no line" on a cell edge, and without a table style PowerPoint draws
 * every unset edge in black; the edges the design leaves open are drawn in the paper colour.
 */
const OPEN = { color: C.bone, width: 0 };
interface Line {
	color: string;
	width: number;
}

function edges(top: Line, bottom: Line): NonNullable<TableCellInput['borders']> {
	return { top, bottom, left: OPEN, right: OPEN };
}

function head(text: string): TableCellInput {
	return {
		text: text.toUpperCase(),
		style: { fontSize: T.kicker, bold: true, color: C.soft, characterSpacing: 1.4 },
		fill: { type: 'none' },
		borders: edges(OPEN, HEAVY),
	};
}

function cell(text: string, bold = false, color: string = C.ink): TableCellInput {
	return { text, style: { fontSize: T.body, bold, color }, fill: { type: 'none' } };
}

/** A body row: hairline below, and above it whatever the row before drew below itself. */
function row(height: number, top: Line, bottom: Line, cells: TableCellInput[]): TableRowInput {
	return { height, cells: cells.map((entry) => ({ ...entry, borders: edges(top, bottom) })) };
}

function results(deck: Deck): void {
	const slide = deck.slide(LAYOUT.canvas);
	slide.placeholder('title', 'What the work did.');
	numbered(slide);
	const rows: TableRowInput[] = [
		{
			height: 40,
			cells: ['Project', 'Discipline', 'Year', 'Scope', 'Weeks', 'Result, first year'].map(head),
		},
		...CASES.map((study, index) =>
			row(80, index === 0 ? HEAVY : HAIRLINE, index === CASES.length - 1 ? HEAVY : HAIRLINE, [
				{
					runs: [
						{
							text: `${study.number}   `,
							style: { fontSize: pt(T.body), bold: true, color: C.vermilion },
						},
						{
							text: study.client,
							style: { fontSize: pt(T.body + 1), bold: true, fontFamily: FONTS.display },
						},
					],
					fill: { type: 'none' },
				},
				cell(study.discipline),
				cell(String(study.year)),
				cell(study.scope),
				cell(String(study.weeks)),
				cell(study.result, true),
			]),
		),
		row(64, HEAVY, OPEN, [
			cell('Three projects', true),
			cell(''),
			cell(''),
			cell('2023 to 2025'),
			cell(String(TOTAL_WEEKS), true),
			cell('All three still in use, unchanged', false, C.soft),
		]),
	];
	const table = slide.table(
		'Results table',
		{ style: 'none', columnWidths: [29, 17, 7, 20, 7, 20], rows },
		{ x: G.left, y: 196, width: G.width, height: 300 },
	);
	// Text sits on the grid line and in the middle of its row.
	for (const line of table.type === 'table' ? (table.tableData?.rows ?? []) : [])
		for (const entry of line.cells) {
			entry.style = { ...entry.style, vAlign: 'middle', marginLeft: 0, marginRight: 16 };
		}
	const source = label(
		slide,
		'Results source',
		'Figures as reported by each client for the twelve months after launch, against the twelve months before.',
		{ x: G.left, y: 566, width: span(8), height: 20, fontSize: pt(T.caption), color: C.soft },
	);
	slide.animate(table, { entrance: 'fadeIn', trigger: 'afterPrevious', durationMs: 700 });
	slide.animate(source, { entrance: 'fadeIn', trigger: 'afterPrevious', durationMs: 500 });
	slide.transition({ type: 'fade', duration: 600 });
	slide.notes(
		`Side by side: ${CASES.map((study) => `${study.client}, ${study.weeks} weeks, ${study.result.toLowerCase()}`).join('; ')}. That is ${TOTAL_WEEKS} weeks of work across three years. The figures come from the clients' own reporting, and all three systems are still in use without changes, which is the result we care about most.`,
	);
}

function process(deck: Deck): void {
	const slide = deck.slide(LAYOUT.canvas);
	slide.placeholder('title', 'Four phases, one team throughout.');
	numbered(slide);
	// The core lays the diagram out itself; `hProcess9` gives four boxes across the full width.
	const data: PptxSmartArtData = applyBuiltinSmartArtLayout(
		{
			layout: 'basicChevronProcess',
			colorScheme: 'monochromatic2',
			style: 'flat',
			nodes: PHASES.map(([name], index) => ({ id: `phase${index + 1}`, text: name })),
		},
		'urn:microsoft.com/office/officeart/2005/8/layout/hProcess9',
	);
	const diagram = slide.smartArt('Process diagram', data, {
		x: G.left,
		y: 184,
		width: G.width,
		height: 150,
	});
	rule(slide, 'Track', G.left, 372, G.width);
	const marker = slide.shape('Marker', 'ellipse', {
		x: G.left - 6,
		y: 366,
		width: 13,
		height: 13,
		fill: { type: 'solid', color: C.vermilion },
		stroke: { color: 'transparent' },
	});
	// The text columns follow the diagram's boxes: 234 px wide, 294 px apart.
	const width = 234;
	PHASES.forEach(([, share, text], index) => {
		const x = G.left + 2 + index * 294.3;
		kicker(slide, `Phase share ${index + 1}`, share, x, 400, width, C.vermilion);
		label(slide, `Phase text ${index + 1}`, text, {
			x,
			y: 424,
			width,
			height: 90,
			fontSize: pt(T.body),
			lineSpacing: 1.35,
		});
	});
	rule(slide, 'Span rule', G.left, 556, G.width);
	const total = label(
		slide,
		'Span',
		[
			[
				{ text: 'First visit to handover: ', style: { color: C.soft } },
				{ text: '11 to 20 weeks', style: { bold: true } },
				{ text: ', as in the three projects shown.', style: { color: C.soft } },
			],
		],
		{ x: G.left, y: 574, width: span(8), height: 24, fontSize: pt(T.body + 1) },
	);
	slide.animate(diagram, {
		entrance: 'wipeIn',
		direction: 'fromLeft',
		trigger: 'afterPrevious',
		durationMs: 900,
	});
	// The marker runs the length of the track: a motion path, relative to the slide width.
	slide.animate(marker, {
		motionPath: `M 0 0 L ${(G.width / 1280).toFixed(4)} 0 E`,
		trigger: 'onClick',
		durationMs: 2400,
	});
	slide.animate(total, { entrance: 'fadeIn', trigger: 'afterPrevious', durationMs: 500 });
	slide.transition({ type: 'push', direction: 'u', duration: 700 });
	slide.notes(
		`Every project runs through the same four phases: ${PHASES.map(([name, share]) => `${name.toLowerCase()} (${share})`).join(', ')}. The people who listen are the people who make and hand over; nothing is passed between departments. Click once and the marker runs the track from first visit to handover.`,
	);
}

function team(deck: Deck): void {
	const slide = deck.slide(LAYOUT.canvas);
	slide.placeholder('title', 'The people who pick up the phone.');
	numbered(slide);
	const width = (G.width - 4 * G.gutter) / 5;
	const height = width * 1.2;
	TEAM.forEach(([name, role], index) => {
		const x = G.left + index * (width + G.gutter);
		const picture = slide.picture(
			`Portrait ${index + 1}`,
			png(`portrait${index + 1}` as PictureName),
			{
				x,
				y: 190,
				width,
				height,
				altText: `${ALT.portrait1}: ${name}`,
			},
		);
		const caption = headline(slide, `Name ${index + 1}`, name, {
			x,
			y: 190 + height + 18,
			width,
			height: 22,
			size: 13.5,
		});
		const line = label(slide, `Role ${index + 1}`, role, {
			x,
			y: 190 + height + 44,
			width,
			height: 34,
			fontSize: pt(T.caption),
			color: C.soft,
			lineSpacing: 1.25,
		});
		// The first portrait follows the transition; the others arrive with it, a beat apart.
		slide.animate(picture, {
			entrance: 'fadeIn',
			trigger: index === 0 ? 'afterPrevious' : 'withPrevious',
			durationMs: 500,
			delayMs: index * 150,
		});
		for (const text of [caption, line])
			slide.animate(text, {
				entrance: 'fadeIn',
				trigger: 'withPrevious',
				durationMs: 500,
				delayMs: index * 150 + 150,
			});
	});
	slide.transition({ type: 'push', direction: 'l', duration: 700 });
	slide.notes(
		`Five of the ${STUDIO.people} of us, and the five you will deal with: ${TEAM.map(([name, role]) => `${name}, ${role.toLowerCase()}`).join('; ')}. The portraits are drawn, not photographed, like everything else in this deck.`,
	);
}

function contact(deck: Deck): void {
	const slide = deck.slide(LAYOUT.dark);
	// A low sun in the corner: a gradient disc with a soft glow behind a smaller ink-blue one.
	slide.shape('Sun', 'ellipse', {
		x: 850,
		y: 150,
		width: 380,
		height: 380,
		fill: {
			type: 'gradient',
			angle: 60,
			stops: [
				{ color: C.wheat, position: 0 },
				{ color: C.vermilion, position: 100 },
			],
		},
		stroke: { color: 'transparent' },
		shadow: { color: C.vermilion, blur: 90, offsetY: 0, opacity: 0.35 },
	});
	slide.shape('Moon', 'ellipse', {
		x: 836,
		y: 372,
		width: 150,
		height: 150,
		fill: { type: 'solid', color: C.blue },
		stroke: { color: 'transparent' },
		shadow: { color: '#000000', blur: 40, offsetY: 14, opacity: 0.35 },
	});
	mark(slide, G.left, G.top, C.bone);
	const thanks = kicker(slide, 'Thanks', 'Thank you for looking', G.left, 250, 400, C.clay);
	slide.placeholder('title', ['Let’s make', 'something', 'that lasts.']);
	numbered(slide);
	const entries: [string, string][] = [
		['Write', STUDIO.email],
		['Visit', STUDIO.address],
		['See more', STUDIO.site],
	];
	const starts = [0, 3, 7];
	const details = entries.flatMap(([name, value], index) => [
		kicker(slide, `Contact kicker ${index + 1}`, name, col(starts[index]!), 552, span(3), C.clay),
		label(slide, `Contact ${index + 1}`, value, {
			x: col(starts[index]!),
			y: 574,
			width: span(4) - 24,
			height: 24,
			fontSize: pt(T.body + 1),
			color: C.bone,
		}),
	]);
	// The thanks leaves on a click, and the contact details take its place.
	slide.animate(thanks, { exit: 'fadeOut', trigger: 'onClick', durationMs: 500 });
	details.forEach((element, index) =>
		slide.animate(element, {
			entrance: 'fadeIn',
			trigger: index === 0 ? 'afterPrevious' : 'withPrevious',
			durationMs: 600,
		}),
	);
	slide.transition({ type: 'fade', duration: 900 });
	slide.notes(
		`That is the work. If something here is close to a problem you have, write to ${STUDIO.email} or visit us at ${STUDIO.address}. We take on about fourteen projects a year, and the first conversation is always with one of the people you just saw.`,
	);
}

export function closingSlides(deck: Deck): void {
	results(deck);
	deck.section('How we work');
	process(deck);
	team(deck);
	deck.section('Contact');
	contact(deck);
}

/** Section "What moved": by mode through the year, by district, and the five busiest routes. */
import type { Deck } from '../kit/deck';
import { districtChart, districtShares, weeklyChart } from './charts';
import {
	DISTRICTS,
	MODE_2024,
	MODE_2025,
	MODES,
	TOTAL_2025,
	TRAM_AFTER,
	TRAM_BEFORE,
	WEEKLY,
	change,
	signed,
	thousands,
} from './data';
import { annotation, card, legend } from './parts';
import { routesSlide } from './slides-routes';
import { C, CONTENT_TOP, MARGIN, MODE_COLORS, frame, label, span } from './style';

const MODE_LEGEND = MODES.map((mode, index) => ({ label: mode, color: MODE_COLORS[index]! }));

function weeklySlide(deck: Deck): void {
	const slide = deck.slide('Chart');
	frame(
		slide,
		{
			kicker: 'What moved',
			headline: 'The tram did the heavy lifting once the Harbour Line extension opened',
			source:
				'Weekly boardings by mode, in thousands. Source: Port Averly Transit automatic passenger counts.',
		},
		4,
	);
	const pills = legend(slide, MODE_LEGEND, MARGIN, CONTENT_TOP);
	label(
		slide,
		'Legend hint',
		'Click Ferry or On-demand for a note',
		{ x: 420, y: CONTENT_TOP + 4, width: 300, height: 18 },
		{ size: 9.5, color: C.slate },
	);
	const chart = slide.chart('Weekly boardings by mode', 'line', weeklyChart(), {
		x: MARGIN - 8,
		y: 228,
		width: 1168,
		height: 388,
	});
	const note = annotation(
		slide,
		'Extension',
		`1 September: the extension opens. Tram boardings go from ${thousands(TRAM_BEFORE)} to ${thousands(TRAM_AFTER)} a week.`,
		{ x: 440, y: 236, width: 372, height: 36 },
		{ x: 862, y: 424 },
	);
	const ferryPeak = Math.max(...WEEKLY[2]!);
	const onDemandShare = ((MODE_2025[3]! / TOTAL_2025) * 100).toFixed(1);
	const notes = [
		{
			name: 'Ferry note',
			x: 120,
			text: `Ferry peaks in July at ${thousands(ferryPeak)} boardings a week, when the harbour is busiest.`,
		},
		{
			name: 'On-demand note',
			x: 440,
			text: `On-demand grew ${signed(change(MODE_2025[3]!, MODE_2024[3]!))}, the fastest of any mode, yet is still ${onDemandShare}% of boardings.`,
		},
	].map((entry) => {
		const box = { x: entry.x, y: 362, width: 296, height: 64 };
		return [
			card(slide, `${entry.name} card`, box),
			label(
				slide,
				entry.name,
				entry.text,
				{ x: box.x + 14, y: box.y + 12, width: box.width - 28, height: box.height - 20 },
				{ size: 10.5, color: C.ink, leading: 1.3 },
			),
		];
	});
	slide.animate(chart, { entrance: 'fadeIn', trigger: 'afterPrevious', durationMs: 800 });
	note.forEach((part, index) =>
		slide.animate(part, {
			entrance: 'fadeIn',
			trigger: index === 0 ? 'afterPrevious' : 'withPrevious',
			durationMs: 500,
		}),
	);
	// Each note belongs to its legend pill: clicking the pill reveals it.
	[pills[2]!, pills[3]!].forEach((pill, index) =>
		notes[index]!.forEach((part) =>
			slide.animate(part, {
				entrance: 'fadeIn',
				trigger: 'onShapeClick',
				triggerShapeId: pill.id,
				durationMs: 400,
			}),
		),
	);
	slide.transition({ type: 'wipe', direction: 'r', duration: 700 });
	slide.notes(
		`Four modes, week by week. Bus is flat and seasonal. The story is the coral line: tram boardings averaged ${thousands(TRAM_BEFORE)} a week until the Harbour Line extension opened on 1 September, and ${thousands(TRAM_AFTER)} a week afterwards, a gain of ${signed(change(TRAM_AFTER, TRAM_BEFORE))}. Clicking the Ferry or On-demand pill in the legend brings up a note on that mode.`,
	);
}

function districtSlide(deck: Deck): void {
	const shares = districtShares();
	const share = (district: string, mode: string) =>
		Math.round(
			shares[DISTRICTS.indexOf(district)]![MODES.indexOf(mode as (typeof MODES)[number])]!,
		);
	const harbourOffRoad = share('Harbourside', 'Tram') + share('Harbourside', 'Ferry');
	const slide = deck.slide('Chart and takeaways');
	frame(
		slide,
		{
			kicker: 'Where it moved',
			headline: 'In Harbourside the tram now carries more people than the bus',
			source:
				'Share of 2025 boardings by mode, by district of boarding. Source: Port Averly Transit.',
		},
		5,
	);
	slide.placeholder(
		'body',
		[
			[
				{ text: 'Harbourside. ', style: { bold: true } },
				{
					text: `Tram ${share('Harbourside', 'Tram')}%, bus ${share('Harbourside', 'Bus')}%. With the ferry, ${harbourOffRoad}% of boardings are off the road.`,
				},
			],
			[
				{ text: 'Saltmeadow. ', style: { bold: true } },
				{
					text: `No tram yet, so on-demand carries ${share('Saltmeadow', 'On-demand')}% of boardings: one trip in seven.`,
				},
			],
			[
				{ text: 'Northgate and Marrow Hill. ', style: { bold: true } },
				{
					text: `Still bus country, at ${share('Northgate', 'Bus')}% and ${share('Marrow Hill', 'Bus')}% of boardings.`,
				},
			],
		],
		13,
	);
	legend(slide, MODE_LEGEND, MARGIN, 204);
	slide.chart('Mode share by district', 'bar', districtChart(), {
		x: MARGIN - 8,
		y: 240,
		width: span(0, 8).width + 16,
		height: 372,
	});
	slide.effects.push(
		{ target: 'Mode share by district', kind: 'wipe', start: 'after', duration: 900 },
		{ target: 'Body 5', paragraph: 0, kind: 'rise', start: 'click' },
		{ target: 'Body 5', paragraph: 1, kind: 'rise', start: 'click' },
		{ target: 'Body 5', paragraph: 2, kind: 'rise', start: 'click' },
	);
	slide.transition({ type: 'push', direction: 'u', duration: 600 });
	slide.notes(
		`Each bar is one district, split by mode. Three clicks bring in the three readings on the right. Harbourside is the first district where the tram outcarries the bus. Saltmeadow has no rail, which is why on-demand does so much work there. Northgate and Marrow Hill depend on the bus, and that is where the Frequent Network matters most.`,
	);
}

export function movedSlides(deck: Deck): void {
	deck.section('What moved');
	weeklySlide(deck);
	districtSlide(deck);
	routesSlide(deck);
}

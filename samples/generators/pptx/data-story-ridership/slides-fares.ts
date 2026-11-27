/** The fare evidence: a doughnut of fare products with a directly labelled list beside it. */
import type { Deck } from '../kit/deck';
import { pt } from '../kit/geometry';
import { FARE_COLORS, fareChart } from './charts';
import { FARES, SINGLE_TICKET_2024 } from './data';
import { C, MARGIN, frame, label, rule, span } from './style';

export function faresSlide(deck: Deck): void {
	const slide = deck.slide('Chart');
	frame(
		slide,
		{
			kicker: 'Why: fares',
			headline: 'Capped tap-and-go became the most used fare in its first year',
			source:
				'Share of 2025 boardings by fare product. Source: Port Averly Transit fare system, all modes.',
		},
		9,
	);
	const chart = slide.chart('Fare products', 'doughnut', fareChart(), {
		x: MARGIN - 4,
		y: 200,
		width: 408,
		height: 408,
	});
	const figure = label(
		slide,
		'Fare figure',
		`${FARES[0]!.share}%`,
		{ x: MARGIN + 100, y: 362, width: 200, height: 60 },
		{
			size: 46,
			display: true,
			bold: true,
			color: C.coral,
			align: 'center',
			tracking: -1,
			leading: 1,
		},
	);
	label(
		slide,
		'Fare figure caption',
		'capped tap-and-go',
		{ x: MARGIN + 100, y: 424, width: 200, height: 20 },
		{ size: 11, color: C.slate, align: 'center' },
	);
	const list = span(6, 6);
	const track = { x: list.x + 220, width: 250 };
	FARES.forEach((fare, index) => {
		const y = 220 + index * 52;
		const color = FARE_COLORS[index]!;
		slide.shape(`Fare dot ${index + 1}`, 'ellipse', {
			x: list.x,
			y: y + 5,
			width: 12,
			height: 12,
			fill: { type: 'solid', color },
			stroke: { color: 'transparent' },
		});
		label(
			slide,
			`Fare name ${index + 1}`,
			fare.name,
			{ x: list.x + 26, y, width: 190, height: 24 },
			{
				size: 13,
				bold: index === 0,
			},
		);
		slide.shape(`Fare track ${index + 1}`, 'roundRect', {
			...track,
			y: y + 7,
			height: 8,
			adjustments: { adj: 50000 },
			fill: { type: 'solid', color: C.sand },
			stroke: { color: 'transparent' },
		});
		slide.shape(`Fare bar ${index + 1}`, 'roundRect', {
			x: track.x,
			y: y + 7,
			width: (track.width * fare.share) / 40,
			height: 8,
			adjustments: { adj: 50000 },
			fill: { type: 'solid', color },
			stroke: { color: 'transparent' },
		});
		label(
			slide,
			`Fare share ${index + 1}`,
			`${fare.share}%`,
			{ x: list.x + list.width - 70, y: y - 2, width: 70, height: 26 },
			{ size: 15, display: true, bold: true, align: 'right', color: index === 0 ? C.coral : C.ink },
		);
		if (index < FARES.length - 1)
			rule(slide, `Fare rule ${index + 1}`, { x: list.x, y: y + 37, width: list.width, height: 1 });
	});
	const pill = (name: string, text: string, y: number, fill: string, color: string) =>
		slide.shape(name, 'roundRect', {
			x: list.x,
			y,
			width: list.width,
			height: 44,
			adjustments: { adj: 18000 },
			fill: { type: 'solid', color: fill },
			stroke: { color: 'transparent' },
			text,
			textStyle: { fontSize: pt(12), color, alignment: 'left', verticalAlignment: 'middle' },
			insets: { left: 20, top: 0, right: 20, bottom: 0 },
		});
	const before = pill(
		'Single tickets 2024',
		`A year earlier, single paper tickets were ${SINGLE_TICKET_2024}% of boardings.`,
		496,
		C.sand,
		C.ink,
	);
	const after = pill(
		'Single tickets 2025',
		`Now they are ${FARES[3]!.share}%. Queues at ticket machines fell with them.`,
		548,
		C.ink,
		C.paper,
	);
	slide.animate(chart, { entrance: 'wheelIn', trigger: 'afterPrevious', durationMs: 900 });
	slide.animate(figure, { entrance: 'zoomIn', trigger: 'afterPrevious', durationMs: 400 });
	slide.animate(figure, { emphasis: 'growShrink', trigger: 'afterPrevious', durationMs: 600 });
	slide.animate(before, { exit: 'fadeOut', trigger: 'onClick', durationMs: 400 });
	slide.animate(after, { entrance: 'fadeIn', trigger: 'withPrevious', durationMs: 500 });
	slide.transition({ type: 'fade', duration: 600 });
	slide.notes(
		`Fares first. Capped tap-and-go launched on 3 March and by year end accounted for ${FARES[0]!.share}% of all boardings, ahead of the monthly pass at ${FARES[1]!.share}%. Click once: single paper tickets fell from ${SINGLE_TICKET_2024}% of boardings in 2024 to ${FARES[3]!.share}%. The five shares add up to 100%.`,
	);
}

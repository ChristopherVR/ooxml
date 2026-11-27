/** Slides 5 and 6: how Aerie works (SmartArt) and what is inside the pod (a table). */
import type {
	TableCellInput,
	TextStyleInput,
} from '../../../../src/core/pptx/core/builders/sdk/index';
import type { PptxSmartArtData } from '../../../../src/core/pptx/core/types/smart-art';
import { applyBuiltinSmartArtLayout } from '../../../../src/core/pptx/smartart-layouts/index';
import { FONTS } from '../../lib/fiction';
import type { Deck } from '../kit/deck';
import { card, head, tx } from './parts';
import { C, G, T, body, col, display, label } from './style';
import { LAYOUT } from './template';

const STEPS = [
	['Sense', 'Every 2.5 seconds', 'Eleven sensors read the air above your own roof.'],
	['Learn', 'In 14 days', 'The model maps how your street differs from the region.'],
	['Forecast', 'Every 10 minutes', 'A fresh 48-hour forecast for your address.'],
	['Alert', '22 minutes ahead', 'A warning on your phone before the weather turns.'],
] as const;

/** The chevrons warm from instrument blue to the amber of an alert. */
export const STEP_STYLE = {
	fills: ['#1F2C4A', '#28416C', '#32598F', C.amber],
	text: [C.text, C.text, C.text, C.night],
	/** Node text size in points. */
	size: 20,
} as const;

function howItWorks(deck: Deck): void {
	const slide = deck.slide(LAYOUT.titleOnly);
	head(slide, 'How it works', 'A gust on your roof, a warning on your phone');
	const data: PptxSmartArtData = applyBuiltinSmartArtLayout(
		{
			layout: 'basicChevronProcess',
			colorScheme: 'colorful1',
			style: 'flat',
			nodes: STEPS.map(([name], index) => ({
				id: `step${index + 1}`,
				text: name,
				style: {
					fillColor: STEP_STYLE.fills[index]!,
					fontColor: STEP_STYLE.text[index]!,
					bold: true,
				},
			})),
		},
		'urn:microsoft.com/office/officeart/2005/8/layout/chevron1',
	);
	const diagram = slide.smartArt('How Aerie works', data, {
		x: G.left,
		y: 228,
		width: G.width,
		height: 150,
	});
	slide.animate(diagram, {
		entrance: 'wipeIn',
		direction: 'fromLeft',
		trigger: 'afterPrevious',
		durationMs: 1200,
	});
	STEPS.forEach(([, when, text], index) => {
		const { x, width } = col(index * 3, 3);
		const heading = tx(
			slide,
			`Step ${index + 1} when`,
			when.toUpperCase(),
			{ x: x + 20, y: 420, width: width - 20, height: 18 },
			label(C.amber),
		);
		const copy = tx(
			slide,
			`Step ${index + 1} text`,
			text,
			{ x: x + 20, y: 448, width: width - 36, height: 90 },
			body(T.body),
			{ lineSpacing: 1.25 },
		);
		for (const element of [heading, copy])
			slide.animate(element, {
				entrance: 'fadeIn',
				trigger: index === 0 && element === heading ? 'afterPrevious' : 'withPrevious',
				durationMs: 500,
				delayMs: index * 220,
			});
	});
	tx(
		slide,
		'Loop',
		'No hub to wire and nothing to calibrate: the loop starts the moment the mast is up.',
		{ x: G.left, y: 580, width: G.width, height: 30 },
		body(T.lead - 2, C.muted),
	);
	slide
		.transition({ type: 'push', direction: 'l', duration: 700 })
		.notes(
			'Four steps, and the numbers under them are the specification, not a slogan. The pod samples every 2.5 seconds; after fourteen days the model has learned the difference between your street and the regional grid; from then on you get a new 48-hour forecast every ten minutes. This diagram is SmartArt, so the audience copy stays editable.',
		);
}

const SIGNALS = [
	['Temperature', '−40 to 60 °C', '± 0.2 °C'],
	['Humidity', '0 to 100 %', '± 1.8 %'],
	['Pressure', '300 to 1,100 hPa', '± 0.5 hPa'],
	['Wind speed', '0 to 60 m/s', '± 3 %'],
	['Wind direction', '0 to 360°', '± 2°'],
	['Rainfall', '0 to 200 mm/h', '± 4 %'],
] as const;

const FACTS = [
	['Power', 'Solar, with a 21-day reserve'],
	['Radio', 'Wi-Fi and LoRa, 400 m range'],
	['Built for', 'IP66, gusts to 60 m/s'],
] as const;

function cell(
	text: string,
	style: Partial<TextStyleInput>,
	rule: { color: string; width: number },
): TableCellInput {
	return {
		text,
		style,
		fill: { type: 'none' },
		borders: {
			top: { color: C.night, width: 0 },
			left: { color: C.night, width: 0 },
			right: { color: C.night, width: 0 },
			bottom: rule,
		},
	};
}

function specifications(deck: Deck): void {
	const slide = deck.slide(LAYOUT.titleOnly);
	head(slide, 'Inside the pod', 'Eleven signals, sampled every 2.5 seconds');
	const headStyle: Partial<TextStyleInput> = {
		fontFamily: FONTS.display,
		fontSize: 9.5,
		bold: true,
		color: C.muted,
		characterSpacing: 2,
	};
	const hairline = { color: C.line, width: 1 };
	const table = slide.table(
		'Sensor specifications',
		{
			style: 'none',
			columnWidths: [0.36, 0.36, 0.28],
			rows: [
				{
					height: 40,
					cells: ['SIGNAL', 'RANGE', 'ACCURACY'].map((text) =>
						cell(text, headStyle, { color: C.amber, width: 2 }),
					),
				},
				...SIGNALS.map(([signal, range, accuracy]) => ({
					height: 48,
					cells: [
						cell(
							signal,
							{ fontFamily: FONTS.display, fontSize: 14, bold: true, color: C.text },
							hairline,
						),
						cell(range, { fontFamily: FONTS.body, fontSize: 13, color: C.soft }, hairline),
						cell(accuracy, { fontFamily: FONTS.body, fontSize: 13, color: C.amber }, hairline),
					],
				})),
			],
		},
		{ x: G.left, y: 204, width: 720, height: 328 },
	);
	if (table.type === 'table')
		for (const row of table.tableData?.rows ?? [])
			for (const item of row.cells)
				item.style = { ...item.style, vAlign: 'middle', marginLeft: 4, marginRight: 12 };
	tx(
		slide,
		'Also on board',
		'Also on board: dew point, UV index, light, lightning distance and PM2.5.',
		{ x: G.left, y: 556, width: 720, height: 22 },
		body(T.small, C.muted),
	);
	FACTS.forEach(([name, value], index) => {
		const top = 204 + index * 116;
		const panel = card(slide, `Fact ${index + 1}`, { x: 848, y: top, width: 352, height: 100 });
		const heading = tx(
			slide,
			`Fact ${index + 1} label`,
			name.toUpperCase(),
			{ x: 872, y: top + 24, width: 304, height: 18 },
			label(C.sky),
		);
		const copy = tx(
			slide,
			`Fact ${index + 1} value`,
			value,
			{ x: 872, y: top + 50, width: 304, height: 28 },
			display(14.5),
		);
		for (const [at, element] of [panel, heading, copy].entries())
			slide.animate(element, {
				entrance: 'fadeIn',
				trigger: at === 0 ? 'onClick' : 'withPrevious',
				durationMs: 400,
			});
	});
	slide
		.transition({ type: 'wipe', direction: 'l', duration: 700 })
		.notes(
			'For the engineers in the room: six of the eleven signals with their range and accuracy, and the other five listed under the table. The accuracy column is what the reference laboratory at the Saltmarsh Institute measured, not what the sensor makers quote. Click three times for power, radio and weatherproofing.',
		);
}

export function productSlides(deck: Deck): void {
	howItWorks(deck);
	specifications(deck);
}

/** What the issue says: the masthead, the lead story, the diary, the plot notes and the sowing guide. */
import type { Block, Paragraph } from '../../../../src/core/docx/index';
import { rich, t, type Doc } from '../kit/builder';
import { cm, pxFromCm, solidFill, textBox } from '../kit/drawing';
import { halfPoints, table, tw } from '../kit/tables';
import { bedsPlan, masthead, sunflower } from './art';
import { GARDEN, TEXT_WIDTH } from './styles';

export const ISSUE = 'Issue 41  ·  Early spring 2026';
export const NAME = 'The Garden Gate';
export const GARDEN_NAME = 'Larkspur Heights Community Garden';
/** Top margin, and the height of the masthead band from the top of the page, in cm. */
export const TOP = 2;
const BAND = 5.76;
export const GUTTER = 454;

/** The masthead: the band behind the text and the name of the newsletter in a box over it. */
export function mastheadBlock(doc: Doc): Paragraph {
	const band = doc.picture(masthead(), {
		name: 'masthead',
		widthPx: pxFromCm(21),
		heightPx: pxFromCm(BAND),
		alt: 'A band of leaves and flower heads in greens, yellow and violet.',
		title: 'Masthead',
		placement: {
			wrap: 'none',
			behindText: true,
			relativeFrom: 'page',
			offsetXPx: 0,
			relativeFromV: 'page',
			offsetYPx: 0,
		},
	});
	const title = textBox(doc, {
		name: 'Masthead text',
		alt: `${NAME}: news from the ${GARDEN_NAME}`,
		width: cm(14),
		height: cm(4),
		x: { from: 'page', offset: cm(1.8) },
		y: { from: 'page', offset: cm(1.05) },
		wrap: 'none',
		fillXml: '<a:noFill/>',
		insets: [0, 0, 0, 0],
		anchor: 'ctr',
		blocks: [
			doc.p('IssueLine', ISSUE, { spacingAfterTwips: tw(160) }),
			doc.p('Title', NAME),
			doc.p('Subtitle', `News from the ${GARDEN_NAME}`, { spacingBeforeTwips: tw(120) }),
		],
	});
	// The paragraph that anchors both reserves the height of the band below the top margin.
	return doc.p('Normal', [band, title], {
		lineSpacingTwips: tw(((BAND + 0.75 - TOP) / 2.54) * 1440),
		lineSpacingRule: 'exact',
		spacingAfterTwips: tw(0),
	});
}

export function lead(doc: Doc): { head: Block[]; story: Block[] } {
	const beds = doc.picture(bedsPlan(), {
		name: 'east-field',
		widthPx: pxFromCm(7.2),
		heightPx: pxFromCm(4.93),
		alt: 'Plan of the east field: twelve raised beds in two rows either side of a path, a hedge along the top.',
		title: 'The east field',
		placement: {
			wrap: 'square',
			wrapText: 'left',
			align: 'right',
			relativeFrom: 'margin',
			relativeFromV: 'paragraph',
			offsetYPx: pxFromCm(0.15),
		},
	});
	const quote = textBox(doc, {
		name: 'Pull quote',
		alt: 'Pull quote from Esi Agyeman',
		width: cm(6.6),
		height: cm(4.1),
		x: { from: 'page', align: 'center' },
		y: { from: 'page', offset: cm(14.4) },
		wrap: 'square',
		distance: [cm(0.2), cm(0.3), cm(0.2), cm(0.3)],
		fillXml: solidFill('accent4'),
		preset: 'roundRect',
		insets: [cm(0.45), cm(0.35), cm(0.45), cm(0.35)],
		anchor: 'ctr',
		blocks: [
			doc.p(
				'PullQuote',
				'“We found rubble, a great deal of clay and a cast-iron bath. The bath is now a pond.”',
			),
			doc.p('PullQuoteSource', 'Esi Agyeman, chair'),
		],
	});
	const head: Block[] = [
		doc.p('Heading1', [beds, t('Forty new beds, one very long hedge')], {
			indentRightTwips: tw(4300),
		}),
		doc.p(
			'Standfirst',
			'The east field is open at last. After two winters of digging, barrowing and tea, the garden has almost doubled in size, and there is a bed waiting for every name on the list.',
			{ indentRightTwips: tw(4300) },
		),
		doc.p('Byline', 'By Calla Whitlow, plot 17', { spacingAfterTwips: tw(200) }),
	];
	const story: Block[] = [
		// The drop cap is a framed paragraph of its own; the text it belongs to follows.
		doc.p('DropCap', [t('O', { fontSize: 51, positionHalfPoints: halfPoints(-8) })], {
			dropCap: { style: 'drop', lines: 3, distanceTwips: tw(60) },
			lineSpacingTwips: tw(880),
			lineSpacingRule: 'exact',
			spacingAfterTwips: tw(0),
		}),
		doc.p('Normal', [
			quote,
			...rich(
				'n the first Saturday of March, Esi Agyeman cut a length of baler twine with her grandmother’s secateurs and declared the east field open. Behind her stood forty raised beds, a hawthorn hedge 112 metres long, and most of the eighty-odd volunteers who built them.',
			),
		]),
		doc.p(
			'Normal',
			'The field was a car park until 2023, when the council agreed to lease it to the association for a peppercorn rent. Getting from tarmac to topsoil took two winters. “We found rubble, a great deal of clay and a cast-iron bath,” says Esi. It took 46 tonnes of compost from the Port Averly green-waste scheme to turn that into something a carrot would tolerate.',
		),
		doc.p(
			'Normal',
			'Every bed is 1.2 by 3 metres, built from larch boards milled at Thistledown and set out so that a wheelchair or a loaded barrow can pass on every side. Six are raised to waist height. A new rainwater main runs the length of the field, fed from the roof of the tool shed.',
		),
		doc.p(
			'Normal',
			rich(
				'The waiting list, which stood at 61 households in the autumn, was cleared in a single evening. Thirty-four beds have gone to new members. The other six are shared: two for Larkspur Primary, two for the Thursday lunch club, and two kept as a nursery for the {sale}.',
				{
					// A direct theme colour on a run: the accent, with its RGB as the stored fallback.
					sale: t('plant sale on 9 May', {
						bold: true,
						color: `#${GARDEN.terracotta}`,
						colorTheme: { token: 'accent2' },
					}),
				},
			),
		),
		doc.p(
			'Normal',
			'The hedge is the long game. Its 560 whips, five to the metre, went in over three frosty weekends, and it will be four or five years before it is thick enough to shelter the beds from the east wind. By then, with luck, the blackbirds will have found it.',
		),
		doc.p(
			'Normal',
			'None of it happened by itself. Thanks are owed to the Thistledown sawmill, to the Port Averly green-waste crew, and to everyone who turned up with a spade when the forecast said to stay in bed.',
			{ spacingAfterTwips: tw(220) },
		),
		doc.p('SidebarHeading', 'In this issue'),
		doc.p('Sidebar', [t('Forty new beds'), t('\t1')]),
		doc.p('Sidebar', [t('Dates for the diary'), t('\t2')]),
		doc.p('Sidebar', [t('Plot notes'), t('\t2')]),
		doc.p('Sidebar', [t('What to sow now'), t('\t2')], { spacingAfterTwips: tw(120) }),
		doc.p(
			'Sidebar',
			rich(
				'**Garden hours.** Open daily from dawn to dusk. The tool shed is staffed on Saturdays from 9 am to 1 pm.',
			),
		),
	];
	return { head, story };
}

const EVENTS: [string, string, string, string][] = [
	['Sat 14 March', 'Seed swap and potato day', 'Tool shed', 'Noor Haddad'],
	['Thu 19 March', 'Committee meeting, 7 pm', 'Larkspur Primary hall', 'Esi Agyeman'],
	['Sat 28 March', 'Hedge mulching work party', 'East field', 'Marek Zielonka'],
	['Sun 5 April', 'Easter egg trail for the under-tens', 'Orchard', 'Beatriz Almada'],
	['Sat 18 April', 'Compost workshop: hot heaps', 'Compost bays', 'Cormac Delaney'],
	['Sat 9 May', 'Plant sale, 10 am to 2 pm', 'Main gate', 'Linnea Sorensen'],
	['Sat 23 May', 'Pond dipping and newt count', 'The bath', 'Yuki Tanabe'],
];

export function diary(doc: Doc): Block[] {
	return [
		doc.p('Heading2', 'Dates for the diary', {
			pageBreakBefore: true,
			spacingBeforeTwips: tw(120),
		}),
		doc.p(
			'Normal',
			'Everything is at the garden unless it says otherwise. Bring gloves; we will bring the kettle.',
		),
		table(doc, {
			style: 'GardenTable',
			textStyle: 'TableText',
			grid: [1750, TEXT_WIDTH - 1750 - 2500 - 2100, 2500, 2100],
			look: {
				firstRow: true,
				lastRow: false,
				firstColumn: true,
				lastColumn: false,
				noHBand: false,
				noVBand: true,
			},
			cantSplit: true,
			rows: [['When', 'What', 'Where', 'Who to ask'], ...EVENTS],
		}),
	];
}

export const NOTES: [string, string][] = [
	[
		'Compost corner',
		'The three hot bays reached 62 °C in February, a garden record. Greens and browns in equal layers, please, and no bindweed: it survives everything except the council incinerator.',
	],
	[
		'Tool library',
		'Forty-three tools are now tagged and hanging on the shed wall. Sign them out in the red book. The long-handled loppers have been missing since November.',
	],
	[
		'Water',
		'The mains tap is for drinking and for the pond only. The six new butts on the shed hold just over 1,300 litres between them, about two dry weeks for the east field.',
	],
	[
		'Little sprouts',
		'Year 3 of Larkspur Primary have sown broad beans in bed E7 and measured them every Friday since. The current leader stands at 14 cm and has been named Gerald.',
	],
	[
		'Pollinator count',
		'Last summer we logged 23 species of bee and hoverfly. To beat it we are sowing phacelia along the new hedge. Counting sheets will be in the shed from May.',
	],
	[
		'Subscriptions',
		'Subs for the year are due by the end of April: €18 a bed, €9 for concessions. Pay Linnea at the shed or by bank transfer. Nobody is turned away for want of it.',
	],
];

export function sowing(doc: Doc): Block[] {
	const flower = doc.picture(sunflower(), {
		name: 'sunflower',
		widthPx: pxFromCm(3.4),
		heightPx: pxFromCm(3.4),
		alt: 'A sunflower head seen face on.',
		title: 'Sunflower',
		placement: {
			wrap: 'tight',
			wrapText: 'left',
			align: 'right',
			relativeFrom: 'column',
			relativeFromV: 'paragraph',
			offsetYPx: pxFromCm(0.1),
		},
	});
	return [
		doc.p('Normal', [
			flower,
			...rich(
				'March is the hinge of the year. The soil is still cold, so sow under cover whatever sulks, and outdoors only what does not mind.',
			),
		]),
		doc.p('ListBullet', rich('**Under glass:** tomatoes, chillies, leeks and sweet peas.')),
		doc.p(
			'ListBullet',
			rich('**Outdoors, under fleece:** broad beans, early peas, spinach and radishes.'),
		),
		doc.p(
			'ListBullet',
			rich('**Wait until April:** carrots, beetroot and anything called a squash.'),
			{
				spacingAfterTwips: tw(110),
			},
		),
		doc.p(
			'Normal',
			'And sunflowers. The tallest-sunflower contest returns, to be judged at the August barbecue; last year Oskar Lindqvist’s 3.4 metre giant took the silver trowel. Seed is free from the shed, one packet a household.',
		),
		doc.p('SidebarHeading', 'Get in touch', { spacingBeforeTwips: tw(200) }),
		doc.p('Sidebar', [
			...rich(
				'**Chair:** Esi Agyeman. **Plots:** Linnea Sorensen. **Newsletter:** Calla Whitlow. ',
			),
			t('Write to '),
			t('gate@larkspurgarden.example', {
				style: 'Hyperlink',
				link: { href: 'mailto:gate@larkspurgarden.example', tooltip: 'Email the garden' },
			}),
			t(' or leave a note in the letterbox on the shed door.'),
		]),
	];
}

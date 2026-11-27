/**
 * Theme, slide master and layouts of "How Port Averly moved": warm paper, ink text, a kicker
 * above every headline, a source line under every chart and a quiet footer.
 */
import type { LayoutSpec, TemplateSpec } from '../kit/template';
import {
	backgroundXml,
	rawPicture,
	rawShape,
	type LevelStyle,
	type PlaceholderSpec,
} from '../kit/xml';
import { coverPanel } from './art';
import { C, MARGIN, span } from './style';

const FULL = span(0, 12);
const HEADLINE = span(0, 9);

const KICKER: LevelStyle = {
	size: 10,
	color: 'accent1',
	font: 'major',
	bold: true,
	tracking: 200,
	caps: true,
};
const SOURCE: LevelStyle = { size: 9.5, color: 'tx2' };
const NUMBER: LevelStyle = { size: 9.5, color: 'tx2', align: 'r' };

const kicker = (id: number, y = 44): PlaceholderSpec => ({
	id,
	name: 'Kicker',
	type: 'body',
	idx: 10,
	x: MARGIN,
	y,
	w: HEADLINE.width,
	h: 20,
	levels: [KICKER],
	prompt: 'Kicker',
});
const headline = (id: number): PlaceholderSpec => ({
	id,
	name: 'Headline',
	type: 'title',
	x: MARGIN,
	y: 68,
	w: HEADLINE.width,
	h: 100,
	prompt: 'A headline that states the finding',
});
const source = (id: number, width = FULL.width): PlaceholderSpec => ({
	id,
	name: 'Source',
	type: 'body',
	idx: 11,
	x: MARGIN,
	y: 626,
	w: width,
	h: 18,
	levels: [SOURCE],
	prompt: 'Source',
});
const number = (id: number): PlaceholderSpec => ({
	id,
	name: 'Slide Number',
	type: 'sldNum',
	idx: 12,
	x: 1136,
	y: 676,
	w: 80,
	h: 18,
	levels: [NUMBER],
});

/** The footer every content slide shares: a hairline, the agency and the report name. */
const FOOTER =
	rawShape({
		id: 40,
		name: 'Footer rule',
		x: MARGIN,
		y: 660,
		w: FULL.width,
		h: 1,
		fill: { solid: C.rule },
	}) +
	rawShape({
		id: 41,
		name: 'Footer agency',
		x: MARGIN,
		y: 676,
		w: 420,
		h: 18,
		runs: [
			{
				text: 'PORT AVERLY TRANSIT',
				size: 8.5,
				color: 'tx1',
				bold: true,
				font: 'major',
				tracking: 160,
			},
			{ text: '     Ridership review 2025', size: 9, color: 'tx2' },
		],
	});

const LAYOUTS: LayoutSpec[] = [
	{
		name: 'Cover',
		type: 'title',
		hideMasterShapes: true,
		images: { rIdCover: 'cover-panel.png' },
		shapes:
			rawPicture({ id: 30, name: 'Cover panel', x: 780, y: 0, w: 500, h: 720, relId: 'rIdCover' }) +
			rawShape({
				id: 31,
				name: 'Cover rule',
				x: MARGIN,
				y: 236,
				w: 56,
				h: 4,
				fill: { solid: C.coral },
			}) +
			rawShape({
				id: 32,
				name: 'Cover agency',
				x: MARGIN,
				y: 64,
				w: 420,
				h: 18,
				runs: [
					{
						text: 'PORT AVERLY TRANSIT',
						size: 10,
						color: 'tx1',
						bold: true,
						font: 'major',
						tracking: 200,
					},
				],
			}),
		placeholders: [
			{
				id: 2,
				name: 'Title',
				type: 'ctrTitle',
				x: MARGIN,
				y: 264,
				w: 640,
				h: 150,
				levels: [
					{ size: 54, color: 'tx1', font: 'major', bold: true, leading: 92, tracking: -100 },
				],
				prompt: 'Report title',
			},
			{
				id: 3,
				name: 'Subtitle',
				type: 'subTitle',
				idx: 1,
				x: MARGIN,
				y: 428,
				w: 600,
				h: 64,
				levels: [{ size: 18, color: 'tx2', leading: 120 }],
				prompt: 'What the report shows',
			},
			{ ...kicker(4, 200), w: 600 },
			{ ...source(5, 600), y: 640, prompt: 'Occasion and date' },
		],
	},
	{
		name: 'Statement',
		type: 'cust',
		placeholders: [kicker(2), headline(3), source(4), number(5)],
	},
	{
		name: 'Chart',
		type: 'obj',
		placeholders: [kicker(2), headline(3), source(4), number(5)],
	},
	{
		name: 'Chart and takeaways',
		type: 'twoObj',
		shapes: rawShape({
			id: 30,
			name: 'Column rule',
			x: span(8, 4).x,
			y: 204,
			w: 1,
			h: 396,
			fill: { solid: C.rule },
		}),
		placeholders: [
			kicker(2),
			headline(3),
			source(4, span(0, 8).width),
			{
				id: 6,
				name: 'Takeaways',
				type: 'body',
				idx: 13,
				x: span(8, 4).x + 24,
				y: 204,
				w: span(8, 4).width - 24,
				h: 396,
				levels: [{ size: 14, color: 'tx1', before: 14, leading: 125 }],
				prompt: 'What to take from the chart',
			},
			number(5),
		],
	},
	{
		name: 'Title only',
		type: 'titleOnly',
		placeholders: [kicker(2), headline(3), number(5)],
	},
	{
		name: 'Closing',
		type: 'cust',
		hideMasterShapes: true,
		background: backgroundXml({ solid: C.ink }),
		shapes:
			rawShape({
				id: 30,
				name: 'Closing rule',
				x: MARGIN,
				y: 660,
				w: FULL.width,
				h: 1,
				fill: { solid: '#3A4656' },
			}) +
			rawShape({
				id: 31,
				name: 'Closing agency',
				x: MARGIN,
				y: 676,
				w: 420,
				h: 18,
				runs: [
					{
						text: 'PORT AVERLY TRANSIT',
						size: 8.5,
						color: C.paper,
						bold: true,
						font: 'major',
						tracking: 160,
					},
					{ text: '     Ridership review 2025', size: 9, color: '#A9B3C1' },
				],
			}),
		placeholders: [
			{ ...kicker(2), levels: [{ ...KICKER, color: C.tealSoft }] },
			{
				...headline(3),
				levels: [
					{ size: 28, color: C.paper, font: 'major', bold: true, leading: 100, tracking: -40 },
				],
			},
			{
				id: 4,
				name: 'Takeaways',
				type: 'body',
				idx: 13,
				x: MARGIN,
				y: 176,
				w: span(0, 7).width,
				h: 320,
				levels: [{ size: 18, color: C.paper, before: 22, leading: 125 }],
				prompt: 'Closing points',
			},
			{ ...number(5), levels: [{ ...NUMBER, color: '#A9B3C1' }] },
		],
	},
];

export const TEMPLATE: TemplateSpec = {
	title: 'How Port Averly moved',
	subject: 'Port Averly Transit ridership review 2025',
	theme: {
		name: 'Port Averly Editorial',
		dark: false,
		colors: {
			dk1: C.ink,
			lt1: C.paper,
			dk2: C.slate,
			lt2: C.sand,
			accent1: C.teal,
			accent2: C.coral,
			accent3: C.ochre,
			accent4: '#3B5B92',
			accent5: C.tealSoft,
			accent6: C.mist,
			hlink: C.teal,
			folHlink: C.slate,
		},
	},
	master: {
		background: backgroundXml({ solid: 'bg1' }),
		shapes: FOOTER,
		placeholders: [
			headline(2),
			{
				id: 3,
				name: 'Body',
				type: 'body',
				idx: 1,
				x: MARGIN,
				y: 196,
				w: FULL.width,
				h: 416,
				prompt: 'Body text',
			},
			number(4),
		],
		titleStyle: { size: 28, color: 'tx1', font: 'major', bold: true, leading: 100, tracking: -40 },
		bodyStyle: [
			{ size: 14, color: 'tx1', before: 8, leading: 120 },
			{ size: 12, color: 'tx2', bullet: '–', margin: 20, indent: -16, before: 4, leading: 120 },
		],
	},
	layouts: LAYOUTS,
	media: { 'cover-panel.png': coverPanel() },
};

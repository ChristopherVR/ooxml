/**
 * The eight layouts of the Aerie keynote and the placeholders they share: a kicker above the
 * title, the title, and the slide number.
 */
import type { LayoutSpec } from '../kit/template';
import { pictureBackgroundXml, rawShape, type LevelStyle, type PlaceholderSpec } from '../kit/xml';
import { G, T } from './style';

export const LAYOUT = {
	cover: 'Cover',
	statement: 'Statement',
	hero: 'Hero',
	titleOnly: 'Title only',
	content: 'Title and content',
	split: 'Content and visual',
	closing: 'Closing',
	blank: 'Blank',
} as const;

export const titleStyle: LevelStyle = {
	size: T.title,
	color: 'tx1',
	font: 'major',
	bold: true,
	leading: 96,
	tracking: -40,
};
const kickerStyle: LevelStyle = {
	size: T.caption,
	color: 'accent1',
	font: 'major',
	bold: true,
	tracking: 220,
	caps: true,
};

/** The kicker above a title (body placeholder 10), the title, and the slide number. */
const kicker = (id: number): PlaceholderSpec => ({
	id,
	name: 'Kicker',
	type: 'body',
	idx: 10,
	x: G.left,
	y: G.top,
	w: 760,
	h: 20,
	levels: [kickerStyle],
	prompt: 'Kicker',
});
export const title = (id: number, w: number = G.width): PlaceholderSpec => ({
	id,
	name: 'Title',
	type: 'title',
	x: G.left,
	y: 92,
	w,
	h: 84,
	prompt: 'Slide title',
});
export const number = (id: number): PlaceholderSpec => ({
	id,
	name: 'Slide Number',
	type: 'sldNum',
	idx: 12,
	x: 1120,
	y: 670,
	w: 80,
	h: 18,
	levels: [{ size: 9, color: 'tx2', align: 'r' }],
});

/** The brand line on the picture layouts, where the master's footer is hidden. */
const brand = (id: number) =>
	rawShape({
		id,
		name: 'Brand',
		x: G.left,
		y: G.top,
		w: 400,
		h: 18,
		runs: [
			{ text: 'NIMBUSWRIGHT', size: 10, color: 'tx1', font: 'major', bold: true, tracking: 380 },
		],
	});

export const layouts: LayoutSpec[] = [
	{
		name: LAYOUT.cover,
		type: 'title',
		hideMasterShapes: true,
		background: pictureBackgroundXml('rIdDusk'),
		images: { rIdDusk: 'dusk.png' },
		shapes: brand(20),
		placeholders: [
			{
				id: 2,
				name: 'Title',
				type: 'ctrTitle',
				x: G.left,
				y: 250,
				w: 700,
				h: 250,
				anchor: 'b',
				levels: [{ ...titleStyle, size: T.hero, leading: 92, tracking: -160 }],
				prompt: 'Deck title',
			},
			{
				id: 3,
				name: 'Subtitle',
				type: 'subTitle',
				idx: 1,
				x: G.left,
				y: 528,
				w: 640,
				h: 56,
				levels: [{ size: 16, color: '#C9D3E6', leading: 130 }],
				prompt: 'Subtitle',
			},
		],
	},
	{
		name: LAYOUT.statement,
		type: 'cust',
		placeholders: [
			kicker(2),
			{
				id: 3,
				name: 'Statement',
				type: 'title',
				x: G.left,
				y: 150,
				w: 900,
				h: 300,
				anchor: 'ctr',
				levels: [{ ...titleStyle, size: T.display, leading: 98, tracking: -110 }],
				prompt: 'One big statement',
			},
			{
				id: 4,
				name: 'Lead',
				type: 'body',
				idx: 11,
				x: G.left,
				y: 480,
				w: 700,
				h: 110,
				levels: [{ size: T.lead, color: 'tx2', leading: 125 }],
				prompt: 'One supporting sentence',
			},
			number(5),
		],
	},
	{
		name: LAYOUT.hero,
		type: 'cust',
		placeholders: [
			kicker(2),
			{
				id: 3,
				name: 'Hero line',
				type: 'title',
				x: G.left,
				y: 190,
				w: 620,
				h: 240,
				anchor: 'ctr',
				levels: [{ ...titleStyle, size: T.hero + 8, leading: 92, tracking: -180 }],
				prompt: 'A few words',
			},
			{
				id: 4,
				name: 'Lead',
				type: 'body',
				idx: 11,
				x: G.left,
				y: 480,
				w: 580,
				h: 110,
				levels: [{ size: T.lead, color: 'tx2', leading: 125 }],
				prompt: 'One supporting sentence',
			},
			number(5),
		],
	},
	{ name: LAYOUT.titleOnly, type: 'titleOnly', placeholders: [kicker(2), title(3), number(4)] },
	{
		name: LAYOUT.content,
		type: 'obj',
		placeholders: [
			kicker(2),
			title(3),
			{ id: 4, name: 'Body', type: 'body', idx: 1, x: G.left, y: G.content, w: G.width, h: 420 },
			number(5),
		],
	},
	{
		name: LAYOUT.split,
		type: 'cust',
		placeholders: [
			kicker(2),
			title(3, 640),
			{ id: 4, name: 'Body', type: 'body', idx: 1, x: G.left, y: 216, w: 600, h: 400 },
			number(5),
		],
	},
	{
		name: LAYOUT.closing,
		type: 'secHead',
		hideMasterShapes: true,
		background: pictureBackgroundXml('rIdDusk'),
		images: { rIdDusk: 'dusk.png' },
		shapes: brand(20),
		placeholders: [
			{
				id: 2,
				name: 'Closing line',
				type: 'title',
				x: G.left,
				y: 180,
				w: 760,
				h: 190,
				anchor: 'b',
				levels: [{ ...titleStyle, size: 88, tracking: -200 }],
				prompt: 'Closing line',
			},
			{
				id: 3,
				name: 'Call to action',
				type: 'body',
				idx: 1,
				x: G.left,
				y: 540,
				w: 900,
				h: 80,
				levels: [{ size: T.lead, color: '#C9D3E6', leading: 125 }],
				prompt: 'Call to action',
			},
		],
	},
	{ name: LAYOUT.blank, type: 'blank', placeholders: [number(2)] },
];

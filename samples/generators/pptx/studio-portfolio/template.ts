/**
 * The portfolio's theme, slide master and layouts. The master carries the paper background and
 * the footer every light slide shares; the dark layout hides it and draws its own.
 */
import type { LayoutSpec, TemplateSpec } from '../kit/template';
import {
	backgroundXml,
	rawShape,
	type LevelStyle,
	type Paint,
	type PlaceholderSpec,
} from '../kit/xml';
import { C, G, T, col, span } from './style';

export const LAYOUT = {
	cover: 'Cover',
	statement: 'Statement',
	canvas: 'Canvas',
	opener: 'Case opener',
	detail: 'Case detail',
	dark: 'Dark feature',
} as const;

const STUDIO = 'Thimblewick Studio';
const DECK = 'Selected work, 2023 to 2025';

/** The footer: a hairline, the studio on the left, the deck name and slide number on the right. */
function footer(ink: Paint, line: Paint, firstId: number): string {
	return (
		rawShape({
			id: firstId,
			name: 'Footer rule',
			x: G.left,
			y: G.footerRule,
			w: G.width,
			h: 1,
			fill: { solid: line },
		}) +
		rawShape({
			id: firstId + 1,
			name: 'Footer studio',
			x: G.left,
			y: G.footerText,
			w: 300,
			h: 14,
			runs: [{ text: STUDIO, size: T.kicker, color: ink, bold: true, tracking: 120, caps: true }],
		}) +
		rawShape({
			id: firstId + 2,
			name: 'Footer deck',
			x: G.right - 360,
			y: G.footerText,
			w: 300,
			h: 14,
			align: 'r',
			runs: [{ text: DECK, size: T.kicker, color: ink }],
		})
	);
}

function number(color: Paint): PlaceholderSpec {
	return {
		id: 5,
		name: 'Slide number',
		type: 'sldNum',
		idx: 12,
		x: G.right - 44,
		y: G.footerText,
		w: 44,
		h: 14,
		levels: [{ size: T.kicker, color, align: 'r', bold: true }],
	};
}

const title = (overrides: Partial<PlaceholderSpec>, level?: LevelStyle): PlaceholderSpec => ({
	id: 2,
	name: 'Title',
	type: 'title',
	x: G.left,
	y: G.top,
	w: span(9),
	h: 84,
	prompt: 'Title',
	...(level ? { levels: [level] } : {}),
	...overrides,
});

const TITLE_STYLE: LevelStyle = {
	size: T.title,
	color: 'tx1',
	font: 'major',
	bold: true,
	leading: 95,
	tracking: -60,
};

const BODY_STYLE: LevelStyle[] = [
	{ size: T.lead, color: 'tx1', leading: 128, before: 10 },
	{
		size: T.body,
		color: 'tx1',
		leading: 135,
		before: 6,
		bullet: '–',
		bulletColor: 'accent1',
		margin: 18,
		indent: -18,
	},
];

const layouts: LayoutSpec[] = [
	{
		name: LAYOUT.cover,
		type: 'title',
		placeholders: [
			{
				id: 2,
				name: 'Title',
				type: 'ctrTitle',
				x: G.left,
				y: 250,
				w: span(7),
				h: 250,
				anchor: 'b',
				prompt: 'Deck title',
				levels: [{ ...TITLE_STYLE, size: T.display, leading: 90, tracking: -140 }],
			},
			{
				id: 3,
				name: 'Subtitle',
				type: 'subTitle',
				idx: 1,
				x: G.left,
				y: 524,
				w: span(6),
				h: 60,
				prompt: 'Subtitle',
				levels: [{ size: T.lead - 4, color: '#6F6A63', leading: 130 }],
			},
			number('tx1'),
		],
	},
	{
		name: LAYOUT.statement,
		type: 'obj',
		placeholders: [
			title({ w: span(6) }),
			{
				id: 3,
				name: 'Body',
				type: 'body',
				idx: 1,
				x: G.left,
				y: 190,
				w: span(6),
				h: 300,
				prompt: 'Statement',
			},
			number('tx1'),
		],
	},
	{ name: LAYOUT.canvas, type: 'titleOnly', placeholders: [title({}), number('tx1')] },
	{
		name: LAYOUT.opener,
		type: 'cust',
		placeholders: [
			title(
				{ y: 396, w: span(3), h: 150, anchor: 't', prompt: 'Client' },
				{ ...TITLE_STYLE, size: T.heading, leading: 100, tracking: -40 },
			),
			number('tx1'),
		],
	},
	{
		name: LAYOUT.detail,
		type: 'cust',
		placeholders: [
			title(
				{ x: col(5), y: 96, w: span(7), h: 44, prompt: 'Project' },
				{ ...TITLE_STYLE, size: T.heading, tracking: -40 },
			),
			{
				id: 3,
				name: 'Brief',
				type: 'body',
				idx: 1,
				x: col(5),
				y: 158,
				w: span(7),
				h: 96,
				prompt: 'The brief',
				levels: [{ size: T.body + 1.5, color: 'tx1', leading: 140 }],
			},
			number('tx1'),
		],
	},
	{
		name: LAYOUT.dark,
		type: 'secHead',
		hideMasterShapes: true,
		background: backgroundXml({ solid: C.night }),
		shapes: footer(C.bone, '#4A4F5E', 20),
		placeholders: [
			title(
				{ y: 300, w: span(8), h: 220, anchor: 'b', prompt: 'Statement' },
				{ ...TITLE_STYLE, size: T.display - 8, color: C.bone, leading: 92, tracking: -120 },
			),
			number(C.bone),
		],
	},
];

export const TEMPLATE: TemplateSpec = {
	title: 'Thimblewick Studio: selected work, 2023 to 2025',
	subject: 'Design portfolio',
	theme: {
		name: 'Thimblewick',
		dark: false,
		colors: {
			dk1: C.ink,
			lt1: C.bone,
			dk2: C.blue,
			lt2: C.paper,
			accent1: C.vermilion,
			accent2: C.blue,
			accent3: C.sage,
			accent4: C.clay,
			accent5: C.wheat,
			accent6: C.moss,
			hlink: C.vermilion,
			folHlink: C.blue,
		},
	},
	master: {
		background: backgroundXml({ solid: 'bg1' }),
		shapes: footer('tx1', C.rule, 20),
		placeholders: [
			title({ name: 'Title placeholder' }),
			{
				id: 3,
				name: 'Body placeholder',
				type: 'body',
				idx: 1,
				x: G.left,
				y: 190,
				w: span(8),
				h: 400,
				prompt: 'Text',
			},
			number('tx1'),
		],
		titleStyle: TITLE_STYLE,
		bodyStyle: BODY_STYLE,
	},
	layouts,
};

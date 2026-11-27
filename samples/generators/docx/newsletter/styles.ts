/**
 * The newsletter's theme and style sheet. Colours and fonts are taken from the theme throughout
 * (`w:themeColor`, `w:themeFill`, theme font attributes), so changing the theme in Word restyles
 * the whole issue.
 */
import { FONTS } from '../../lib/fiction';
import { numIdOf, type ListSpec } from '../kit/parts';
import { GARDEN } from './theme';

export { GARDEN, THEME } from './theme';
import {
	BASE_STYLES,
	tableRegion,
	type Border,
	type ParagraphLook,
	type RunLook,
	type StyleSheet,
} from '../kit/styles';

/** A4 with 1.8 cm side margins: a wide text area for columns. */
export const MARGIN = 1021;
export const TEXT_WIDTH = 11906 - 2 * MARGIN;

/** Theme colours as style properties: the theme name with its RGB as the stored fallback. */
const themed = (theme: string, rgb: string): Pick<RunLook, 'color' | 'themeColor'> => ({
	color: rgb,
	themeColor: theme,
});
const display: RunLook = { themeFont: 'major', font: FONTS.body, bold: true };
const line = (theme: string, rgb: string, size: number, space = 0): Border => ({
	size,
	space,
	color: rgb,
	themeColor: theme,
});

/** The frame every paragraph of a sidebar shares: a cream panel under a leaf-green rule. */
const SIDEBAR_BOX: ParagraphLook = {
	borders: {
		top: line('accent1', GARDEN.leaf, 24, 7),
		left: line('background2', GARDEN.cream, 4, 8),
		bottom: line('accent1', GARDEN.leaf, 6, 7),
		right: line('background2', GARDEN.cream, 4, 8),
	},
	fill: GARDEN.cream,
	themeFill: 'background2',
	left: 170,
	right: 170,
};

const BULLETS: ListSpec = {
	id: 0,
	type: 'hybridMultilevel',
	levels: [
		{
			format: 'bullet',
			text: '•',
			paragraph: { left: 284, hanging: 227 },
			run: { font: FONTS.body, ...themed('accent2', GARDEN.terracotta) },
		},
	],
};
export const LISTS = [BULLETS];
export const BULLET_NUM = numIdOf(BULLETS);

export const STYLE_SHEET: StyleSheet = {
	run: {
		themeFont: 'minor',
		font: FONTS.body,
		size: 9.5,
		...themed('text1', GARDEN.ink),
		lang: 'en-GB',
	},
	paragraph: { after: 110, line: 276 },
	styles: [
		...BASE_STYLES,
		{
			type: 'paragraph',
			id: 'Title',
			name: 'Title',
			basedOn: 'Normal',
			next: 'Normal',
			uiPriority: 10,
			quick: true,
			paragraph: { after: 0, line: 216 },
			run: { ...display, size: 38, spacing: -10, ...themed('background1', GARDEN.white) },
		},
		{
			type: 'paragraph',
			id: 'Subtitle',
			name: 'Subtitle',
			basedOn: 'Normal',
			next: 'Normal',
			uiPriority: 11,
			quick: true,
			paragraph: { after: 0, line: 264 },
			run: { size: 11, ...themed('background2', GARDEN.cream) },
		},
		{
			type: 'paragraph',
			id: 'IssueLine',
			name: 'Issue Line',
			custom: true,
			basedOn: 'Normal',
			uiPriority: 12,
			paragraph: { after: 0, line: 240 },
			run: {
				bold: true,
				caps: true,
				size: 8.5,
				spacing: 30,
				...themed('accent3', GARDEN.marigold),
			},
		},
		{
			type: 'paragraph',
			id: 'Heading1',
			name: 'heading 1',
			basedOn: 'Normal',
			next: 'Normal',
			uiPriority: 9,
			quick: true,
			paragraph: {
				keepNext: true,
				keepLines: true,
				before: 0,
				after: 80,
				line: 228,
				outlineLevel: 0,
			},
			run: { ...display, size: 25, spacing: -6, ...themed('text2', GARDEN.forest) },
		},
		{
			type: 'paragraph',
			id: 'Standfirst',
			name: 'Standfirst',
			custom: true,
			basedOn: 'Normal',
			next: 'Normal',
			uiPriority: 13,
			quick: true,
			paragraph: { after: 200, line: 300 },
			run: { size: 12, ...themed('accent6', GARDEN.bark) },
		},
		{
			type: 'paragraph',
			id: 'Heading2',
			name: 'heading 2',
			basedOn: 'Normal',
			next: 'Normal',
			uiPriority: 9,
			quick: true,
			paragraph: {
				keepNext: true,
				keepLines: true,
				borders: { bottom: line('accent3', GARDEN.marigold, 18, 4) },
				before: 240,
				after: 160,
				line: 240,
				outlineLevel: 1,
			},
			run: { ...display, size: 17, ...themed('text2', GARDEN.forest) },
		},
		{
			type: 'paragraph',
			id: 'Heading3',
			name: 'heading 3',
			basedOn: 'Normal',
			next: 'Normal',
			uiPriority: 9,
			quick: true,
			paragraph: {
				keepNext: true,
				keepLines: true,
				before: 140,
				after: 40,
				line: 240,
				outlineLevel: 2,
			},
			run: { ...display, size: 10.5, ...themed('accent2', GARDEN.terracotta) },
		},
		{
			type: 'paragraph',
			id: 'Byline',
			name: 'Byline',
			custom: true,
			basedOn: 'Normal',
			next: 'Normal',
			uiPriority: 14,
			paragraph: { after: 120 },
			run: { bold: true, caps: true, size: 7.5, spacing: 20, ...themed('accent1', GARDEN.leaf) },
		},
		{
			type: 'paragraph',
			id: 'DropCap',
			name: 'Drop Cap',
			custom: true,
			basedOn: 'Normal',
			uiPriority: 15,
			run: { ...display, ...themed('accent2', GARDEN.terracotta) },
		},
		{
			type: 'paragraph',
			id: 'PullQuote',
			name: 'Pull Quote',
			custom: true,
			basedOn: 'Normal',
			uiPriority: 16,
			quick: true,
			paragraph: { after: 60, line: 300, align: 'center' },
			run: { font: FONTS.serif, italic: true, size: 13, ...themed('background1', GARDEN.white) },
		},
		{
			type: 'paragraph',
			id: 'PullQuoteSource',
			name: 'Pull Quote Source',
			custom: true,
			basedOn: 'Normal',
			uiPriority: 16,
			paragraph: { after: 0, line: 240, align: 'center' },
			run: {
				bold: true,
				caps: true,
				size: 7.5,
				spacing: 20,
				...themed('accent3', GARDEN.marigold),
			},
		},
		{
			type: 'paragraph',
			id: 'SidebarHeading',
			name: 'Sidebar Heading',
			custom: true,
			basedOn: 'Normal',
			next: 'Sidebar',
			uiPriority: 17,
			quick: true,
			// The same borders as Sidebar: Word draws paragraphs with equal borders as one box.
			paragraph: { keepNext: true, ...SIDEBAR_BOX, after: 40, line: 264 },
			run: { ...display, size: 11, ...themed('text2', GARDEN.forest) },
		},
		{
			type: 'paragraph',
			id: 'Sidebar',
			name: 'Sidebar',
			custom: true,
			basedOn: 'Normal',
			uiPriority: 17,
			quick: true,
			paragraph: {
				...SIDEBAR_BOX,
				tabs: [{ pos: 4480, align: 'right', leader: 'dot' }],
				after: 20,
				line: 280,
			},
			run: { size: 9 },
		},
		{
			type: 'paragraph',
			id: 'ListBullet',
			name: 'List Bullet',
			basedOn: 'Normal',
			uiPriority: 36,
			quick: true,
			paragraph: { numbering: { numId: BULLET_NUM, level: 0 }, after: 40, contextualSpacing: true },
		},
		{
			type: 'paragraph',
			id: 'TableText',
			name: 'Table Text',
			custom: true,
			basedOn: 'Normal',
			uiPriority: 41,
			paragraph: { before: 0, after: 0, line: 252 },
			run: { size: 9 },
		},
		{
			type: 'paragraph',
			id: 'Caption',
			name: 'caption',
			basedOn: 'Normal',
			next: 'Normal',
			uiPriority: 35,
			quick: true,
			paragraph: { after: 0, line: 240 },
			run: { italic: true, size: 8, ...themed('accent6', GARDEN.bark) },
		},
		{
			type: 'paragraph',
			id: 'Header',
			name: 'header',
			basedOn: 'Normal',
			uiPriority: 99,
			unhideWhenUsed: true,
			paragraph: { tabs: [{ pos: TEXT_WIDTH, align: 'right' }], after: 0, line: 240 },
			run: { bold: true, caps: true, size: 7.5, spacing: 20, ...themed('accent1', GARDEN.leaf) },
		},
		{
			type: 'paragraph',
			id: 'Footer',
			name: 'footer',
			basedOn: 'Normal',
			uiPriority: 99,
			unhideWhenUsed: true,
			paragraph: {
				borders: { top: line('accent1', GARDEN.leaf, 12, 8) },
				tabs: [{ pos: TEXT_WIDTH, align: 'right' }],
				after: 0,
				line: 252,
			},
			run: { size: 8, ...themed('text2', GARDEN.forest) },
		},
		{
			type: 'character',
			id: 'Hyperlink',
			name: 'Hyperlink',
			basedOn: 'DefaultParagraphFont',
			uiPriority: 99,
			unhideWhenUsed: true,
			run: { ...themed('hyperlink', GARDEN.forest), underline: 'single' },
		},
		{
			type: 'character',
			id: 'Strong',
			name: 'Strong',
			basedOn: 'DefaultParagraphFont',
			uiPriority: 22,
			quick: true,
			run: { bold: true },
		},
		{
			type: 'table',
			id: 'GardenTable',
			name: 'Garden Table',
			custom: true,
			basedOn: 'TableNormal',
			uiPriority: 40,
			tableXml: [
				'<w:tblPr><w:tblStyleRowBandSize w:val="1"/><w:tblStyleColBandSize w:val="1"/>',
				`<w:tblBorders><w:bottom w:val="single" w:sz="12" w:space="0" w:color="${GARDEN.forest}" w:themeColor="text2"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="${GARDEN.cream}" w:themeColor="background2"/></w:tblBorders>`,
				'<w:tblCellMar><w:top w:w="68" w:type="dxa"/><w:left w:w="113" w:type="dxa"/><w:bottom w:w="68" w:type="dxa"/><w:right w:w="113" w:type="dxa"/></w:tblCellMar></w:tblPr>',
				tableRegion('firstRow', {
					run: { bold: true, ...themed('background1', GARDEN.white) },
					fill: GARDEN.forest,
					themeFill: 'text2',
				}),
				tableRegion('firstCol', { run: { bold: true, ...themed('accent2', GARDEN.terracotta) } }),
				tableRegion('band2Horz', { fill: GARDEN.cream, themeFill: 'background2' }),
			].join(''),
		},
	],
};

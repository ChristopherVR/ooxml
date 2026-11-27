/** The report's style sheet: the structural styles here, the text styles in `styles-text.ts`. */
import { FONTS } from '../../lib/fiction';
import {
	BASE_STYLES,
	linked,
	rule,
	tableRegion,
	type StyleSheet,
	type StyleSpec,
} from '../kit/styles';
import { PALETTE } from './data';
import { NUM, TEXT_WIDTH } from './layout';
import { TEXT_STYLES } from './styles-text';

export { LISTS, MARGIN, NUM, TEXT_WIDTH, THEME, WIDE_WIDTH } from './layout';

const heading = (level: 1 | 2 | 3, size: number, color: string, before: number): StyleSpec[] =>
	linked({
		type: 'paragraph',
		id: `Heading${level}`,
		name: `heading ${level}`,
		basedOn: 'Normal',
		next: 'Normal',
		uiPriority: 9,
		quick: true,
		paragraph: {
			keepNext: true,
			keepLines: true,
			numbering: { numId: NUM.headings, level: level - 1 },
			before,
			after: level === 1 ? 200 : 100,
			line: 264,
			outlineLevel: level - 1,
		},
		run: { font: FONTS.display, bold: true, color, size },
	});

const toc = (level: 1 | 2 | 3): StyleSpec => {
	const left = [0, 567, 1247][level - 1]!;
	const numberTab = [567, 1247, 2041][level - 1]!;
	return {
		type: 'paragraph',
		id: `TOC${level}`,
		name: `toc ${level}`,
		basedOn: 'Normal',
		next: 'Normal',
		uiPriority: 39,
		unhideWhenUsed: true,
		paragraph: {
			tabs: [
				{ pos: numberTab, align: 'left' },
				{ pos: TEXT_WIDTH, align: 'right', leader: 'dot' },
			],
			before: level === 1 ? 160 : 0,
			after: 40,
			line: 276,
			left,
		},
		run: {
			font: FONTS.body,
			size: level === 1 ? 10 : 9.5,
			...(level === 1 && { bold: true, color: PALETTE.deep }),
		},
	};
};

const TABLE_STYLE: StyleSpec = {
	type: 'table',
	id: 'SaltmarshTable',
	name: 'Saltmarsh Table',
	custom: true,
	basedOn: 'TableNormal',
	uiPriority: 40,
	paragraph: { before: 0, after: 0, line: 252 },
	run: { font: FONTS.body, size: 9 },
	tableXml: [
		'<w:tblPr><w:tblStyleRowBandSize w:val="1"/><w:tblStyleColBandSize w:val="1"/>',
		`<w:tblBorders><w:top w:val="single" w:sz="12" w:space="0" w:color="${PALETTE.deep}"/><w:bottom w:val="single" w:sz="12" w:space="0" w:color="${PALETTE.deep}"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="${PALETTE.rule}"/></w:tblBorders>`,
		'<w:tblCellMar><w:top w:w="57" w:type="dxa"/><w:left w:w="113" w:type="dxa"/><w:bottom w:w="57" w:type="dxa"/><w:right w:w="113" w:type="dxa"/></w:tblCellMar></w:tblPr>',
		tableRegion('firstRow', {
			run: { bold: true, color: PALETTE.white },
			fill: PALETTE.deep,
		}),
		tableRegion('lastRow', {
			run: { bold: true, color: PALETTE.deep },
			top: rule(PALETTE.deep, 8),
		}),
		tableRegion('firstCol', { run: { bold: true, color: PALETTE.deep } }),
		tableRegion('band1Horz', { fill: PALETTE.white }),
		tableRegion('band2Horz', { fill: PALETTE.foam }),
	].join(''),
};

export const STYLE_SHEET: StyleSheet = {
	run: { font: FONTS.serif, size: 10.5, color: PALETTE.ink, lang: 'en-GB' },
	paragraph: { after: 150, line: 312 },
	styles: [
		...BASE_STYLES,
		...linked({
			type: 'paragraph',
			id: 'Title',
			name: 'Title',
			basedOn: 'Normal',
			next: 'Subtitle',
			uiPriority: 10,
			quick: true,
			paragraph: { after: 120, line: 228, contextualSpacing: true },
			run: { font: FONTS.display, bold: true, color: PALETTE.deep, size: 34, spacing: -10 },
		}),
		...linked({
			type: 'paragraph',
			id: 'Subtitle',
			name: 'Subtitle',
			basedOn: 'Normal',
			next: 'Normal',
			uiPriority: 11,
			quick: true,
			paragraph: { after: 360, line: 264 },
			// Rubik is always set bold: one weight for every display line.
			run: { font: FONTS.display, bold: true, color: PALETTE.teal, size: 15 },
		}),
		...heading(1, 19, PALETTE.deep, 420),
		...heading(2, 13.5, PALETTE.teal, 300),
		...heading(3, 11, PALETTE.ink, 220),
		{
			type: 'paragraph',
			id: 'HeadingUnnumbered',
			name: 'Heading Unnumbered',
			custom: true,
			basedOn: 'Heading1',
			next: 'Normal',
			uiPriority: 9,
			quick: true,
			paragraph: { numbering: { numId: 0, level: 0 }, left: 0, hanging: 0 },
		},
		{
			type: 'paragraph',
			id: 'FrontHeading',
			name: 'Front Heading',
			custom: true,
			basedOn: 'Normal',
			next: 'Normal',
			uiPriority: 9,
			quick: true,
			paragraph: { keepNext: true, before: 0, after: 240, line: 264 },
			run: { font: FONTS.display, bold: true, color: PALETTE.deep, size: 19 },
		},
		toc(1),
		toc(2),
		toc(3),
		{
			type: 'paragraph',
			id: 'TableofFigures',
			name: 'table of figures',
			basedOn: 'Normal',
			next: 'Normal',
			uiPriority: 99,
			unhideWhenUsed: true,
			paragraph: {
				tabs: [{ pos: TEXT_WIDTH, align: 'right', leader: 'dot' }],
				after: 60,
				line: 276,
				left: 1021,
				hanging: 1021,
			},
			run: { font: FONTS.body, size: 9.5 },
		},
		{
			type: 'paragraph',
			id: 'Caption',
			name: 'caption',
			basedOn: 'Normal',
			next: 'Normal',
			uiPriority: 35,
			quick: true,
			paragraph: { keepLines: true, before: 100, after: 240, line: 264 },
			run: { font: FONTS.body, size: 9, color: PALETTE.slate },
		},
		{
			type: 'character',
			id: 'CaptionLabel',
			name: 'Caption Label',
			custom: true,
			basedOn: 'DefaultParagraphFont',
			uiPriority: 35,
			run: { bold: true, color: PALETTE.deep },
		},
		{
			type: 'paragraph',
			id: 'Figure',
			name: 'Figure',
			custom: true,
			basedOn: 'Normal',
			next: 'Caption',
			uiPriority: 34,
			quick: true,
			paragraph: { keepNext: true, before: 200, after: 0, line: 240, align: 'center' },
		},
		...linked({
			type: 'paragraph',
			id: 'Quote',
			name: 'Quote',
			basedOn: 'Normal',
			next: 'Normal',
			uiPriority: 29,
			quick: true,
			paragraph: {
				borders: { left: rule(PALETTE.amber, 24, 12) },
				before: 200,
				after: 240,
				left: 425,
				right: 567,
			},
			run: { italic: true, color: PALETTE.deep, size: 11 },
		}),
		{
			type: 'paragraph',
			id: 'KeyFinding',
			name: 'Key Finding',
			custom: true,
			basedOn: 'Normal',
			next: 'Normal',
			uiPriority: 30,
			quick: true,
			paragraph: {
				keepLines: true,
				borders: { top: rule(PALETTE.teal, 12, 8), bottom: rule(PALETTE.rule, 4, 8) },
				fill: PALETTE.foam,
				before: 240,
				after: 280,
				line: 288,
				left: 170,
				right: 170,
			},
			run: { font: FONTS.body, size: 9.5, color: PALETTE.ink },
		},
		...TEXT_STYLES,
		TABLE_STYLE,
		{
			type: 'table',
			id: 'LayoutGrid',
			name: 'Layout Grid',
			custom: true,
			basedOn: 'TableNormal',
			uiPriority: 45,
			tableXml:
				'<w:tblPr><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="0" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar></w:tblPr>',
		},
	],
};

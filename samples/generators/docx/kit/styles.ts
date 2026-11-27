/**
 * `word/styles.xml`, written by hand. The core writes a styles part for a new document only from
 * its built-in defaults (which name Microsoft fonts), so each sample states its whole style sheet
 * here: document defaults, paragraph, character and table styles, in schema order.
 */
import { W_NS } from './builder';

export interface RunLook {
	/** One family for every script slot, so nothing falls back to a font the file does not name. */
	font?: string;
	/** `major` or `minor`: the theme font, with `font` as the East Asian and complex-script face. */
	themeFont?: 'major' | 'minor';
	bold?: boolean;
	italic?: boolean;
	caps?: boolean;
	smallCaps?: boolean;
	/** `RRGGBB`. */
	color?: string;
	/** A theme colour (`accent1`, `text2`...), with an optional tint or shade byte (`BF`). */
	themeColor?: string;
	themeTint?: string;
	themeShade?: string;
	/** Character spacing in twentieths of a point. */
	spacing?: number;
	/** Points. */
	size?: number;
	underline?: 'single' | 'none';
	vertAlign?: 'superscript' | 'subscript';
	/** Run shading fill, `RRGGBB`. */
	fill?: string;
	lang?: string;
}

/** `w:shd` with a plain fill, or a theme fill (`background2`...) with its RGB as the fallback. */
const shading = (fill: string | undefined, theme: string | undefined): string =>
	fill || theme
		? `<w:shd w:val="clear" w:color="auto" w:fill="${fill ?? 'auto'}"${theme ? ` w:themeFill="${theme}"` : ''}/>`
		: '';

/** `w:rPr` children in CT_RPr order. */
export function runLook(look: RunLook): string {
	const fonts = look.themeFont
		? `<w:rFonts w:asciiTheme="${look.themeFont}HAnsi" w:hAnsiTheme="${look.themeFont}HAnsi"${
				look.font ? ` w:eastAsia="${look.font}" w:cs="${look.font}"` : ''
			}/>`
		: look.font
			? `<w:rFonts w:ascii="${look.font}" w:hAnsi="${look.font}" w:eastAsia="${look.font}" w:cs="${look.font}"/>`
			: '';
	const toggle = (name: string, value: boolean | undefined) =>
		value === undefined ? '' : value ? `<w:${name}/>` : `<w:${name} w:val="0"/>`;
	const color =
		look.themeColor || look.color
			? `<w:color w:val="${look.color ?? 'auto'}"${
					look.themeColor ? ` w:themeColor="${look.themeColor}"` : ''
				}${look.themeTint ? ` w:themeTint="${look.themeTint}"` : ''}${
					look.themeShade ? ` w:themeShade="${look.themeShade}"` : ''
				}/>`
			: '';
	return [
		fonts,
		toggle('b', look.bold),
		toggle('bCs', look.bold),
		toggle('i', look.italic),
		toggle('iCs', look.italic),
		toggle('caps', look.caps),
		toggle('smallCaps', look.smallCaps),
		color,
		look.spacing === undefined ? '' : `<w:spacing w:val="${look.spacing}"/>`,
		look.size === undefined
			? ''
			: `<w:sz w:val="${Math.round(look.size * 2)}"/><w:szCs w:val="${Math.round(look.size * 2)}"/>`,
		look.underline ? `<w:u w:val="${look.underline}"/>` : '',
		look.fill ? `<w:shd w:val="clear" w:color="auto" w:fill="${look.fill}"/>` : '',
		look.vertAlign ? `<w:vertAlign w:val="${look.vertAlign}"/>` : '',
		look.lang ? `<w:lang w:val="${look.lang}" w:eastAsia="${look.lang}" w:bidi="ar-SA"/>` : '',
	].join('');
}

export interface Border {
	/** `single`, `double`, `dotted`... */
	style?: string;
	/** Eighths of a point. */
	size: number;
	/** Points between the line and the text. */
	space?: number;
	color: string;
	/** A theme colour the line follows; `color` is then its RGB fallback. */
	themeColor?: string;
}

const border = (side: string, value: Border | undefined) =>
	value
		? `<w:${side} w:val="${value.style ?? 'single'}" w:sz="${value.size}" w:space="${value.space ?? 0}" w:color="${value.color}"${value.themeColor ? ` w:themeColor="${value.themeColor}"` : ''}/>`
		: '';

export interface TabLook {
	pos: number;
	align: 'left' | 'center' | 'right' | 'decimal';
	leader?: 'dot' | 'underscore' | 'hyphen';
}

export interface ParagraphLook {
	keepNext?: boolean;
	keepLines?: boolean;
	pageBreakBefore?: boolean;
	widowControl?: boolean;
	/** The numbering instance and level the style is bound to. */
	numbering?: { numId: number; level: number };
	borders?: { top?: Border; left?: Border; bottom?: Border; right?: Border; between?: Border };
	/** Paragraph shading fill, `RRGGBB`. */
	fill?: string;
	/** A theme colour for the shading (`background2`...); `fill` is then its RGB fallback. */
	themeFill?: string;
	tabs?: TabLook[];
	suppressAutoHyphens?: boolean;
	/** Twentieths of a point. */
	before?: number;
	after?: number;
	/** `auto` line spacing in 240ths of a line, or twips with `lineRule`. */
	line?: number;
	lineRule?: 'auto' | 'exact' | 'atLeast';
	left?: number;
	right?: number;
	firstLine?: number;
	hanging?: number;
	contextualSpacing?: boolean;
	align?: 'left' | 'center' | 'right' | 'both';
	/** 0 to 8, or 9 for body text. */
	outlineLevel?: number;
}

/** `w:pPr` children in CT_PPrBase order. */
export function paragraphLook(look: ParagraphLook): string {
	const toggle = (name: string, value: boolean | undefined) =>
		value === undefined ? '' : value ? `<w:${name}/>` : `<w:${name} w:val="0"/>`;
	const borders = look.borders
		? `<w:pBdr>${(['top', 'left', 'bottom', 'right', 'between'] as const)
				.map((side) => border(side, look.borders?.[side]))
				.join('')}</w:pBdr>`
		: '';
	const spacing =
		look.before !== undefined || look.after !== undefined || look.line !== undefined
			? `<w:spacing${look.before === undefined ? '' : ` w:before="${look.before}"`}${
					look.after === undefined ? '' : ` w:after="${look.after}"`
				}${
					look.line === undefined
						? ''
						: ` w:line="${look.line}" w:lineRule="${look.lineRule ?? 'auto'}"`
				}/>`
			: '';
	const indent =
		look.left !== undefined ||
		look.right !== undefined ||
		look.firstLine !== undefined ||
		look.hanging !== undefined
			? `<w:ind${look.left === undefined ? '' : ` w:left="${look.left}"`}${
					look.right === undefined ? '' : ` w:right="${look.right}"`
				}${look.firstLine === undefined ? '' : ` w:firstLine="${look.firstLine}"`}${
					look.hanging === undefined ? '' : ` w:hanging="${look.hanging}"`
				}/>`
			: '';
	return [
		toggle('keepNext', look.keepNext),
		toggle('keepLines', look.keepLines),
		toggle('pageBreakBefore', look.pageBreakBefore),
		toggle('widowControl', look.widowControl),
		look.numbering
			? `<w:numPr><w:ilvl w:val="${look.numbering.level}"/><w:numId w:val="${look.numbering.numId}"/></w:numPr>`
			: '',
		borders,
		shading(look.fill, look.themeFill),
		look.tabs
			? `<w:tabs>${look.tabs
					.map(
						(tab) =>
							`<w:tab w:val="${tab.align}"${tab.leader ? ` w:leader="${tab.leader}"` : ''} w:pos="${tab.pos}"/>`,
					)
					.join('')}</w:tabs>`
			: '',
		toggle('suppressAutoHyphens', look.suppressAutoHyphens),
		spacing,
		indent,
		toggle('contextualSpacing', look.contextualSpacing),
		look.align ? `<w:jc w:val="${look.align}"/>` : '',
		look.outlineLevel === undefined ? '' : `<w:outlineLvl w:val="${look.outlineLevel}"/>`,
	].join('');
}

export interface StyleSpec {
	type: 'paragraph' | 'character' | 'table' | 'numbering';
	id: string;
	/** The name Word shows; built-in styles must use Word's own names (`heading 1`, `toc 1`). */
	name: string;
	isDefault?: boolean;
	/** A style Word does not ship: marked `w:customStyle`. */
	custom?: boolean;
	basedOn?: string;
	next?: string;
	link?: string;
	uiPriority?: number;
	semiHidden?: boolean;
	unhideWhenUsed?: boolean;
	quick?: boolean;
	paragraph?: ParagraphLook;
	run?: RunLook;
	/** Table styles: `w:tblPr`, `w:trPr`, `w:tcPr` and `w:tblStylePr` markup, in that order. */
	tableXml?: string;
}

function styleXml(style: StyleSpec): string {
	const paragraph = style.paragraph ? paragraphLook(style.paragraph) : '';
	const run = style.run ? runLook(style.run) : '';
	return [
		`<w:style w:type="${style.type}"${style.isDefault ? ' w:default="1"' : ''}${
			style.custom ? ' w:customStyle="1"' : ''
		} w:styleId="${style.id}">`,
		`<w:name w:val="${style.name}"/>`,
		style.basedOn ? `<w:basedOn w:val="${style.basedOn}"/>` : '',
		style.next ? `<w:next w:val="${style.next}"/>` : '',
		style.link ? `<w:link w:val="${style.link}"/>` : '',
		style.uiPriority === undefined ? '' : `<w:uiPriority w:val="${style.uiPriority}"/>`,
		style.semiHidden ? '<w:semiHidden/>' : '',
		style.unhideWhenUsed ? '<w:unhideWhenUsed/>' : '',
		style.quick ? '<w:qFormat/>' : '',
		paragraph ? `<w:pPr>${paragraph}</w:pPr>` : '',
		run ? `<w:rPr>${run}</w:rPr>` : '',
		style.tableXml ?? '',
		'</w:style>',
	].join('');
}

export interface StyleSheet {
	/** Document default run properties (font, size, language). */
	run: RunLook;
	/** Document default paragraph properties. */
	paragraph: ParagraphLook;
	styles: StyleSpec[];
}

/** The four default styles every Word style sheet starts from. */
export const BASE_STYLES: StyleSpec[] = [
	{ type: 'paragraph', id: 'Normal', name: 'Normal', isDefault: true, quick: true },
	{
		type: 'character',
		id: 'DefaultParagraphFont',
		name: 'Default Paragraph Font',
		isDefault: true,
		uiPriority: 1,
		semiHidden: true,
		unhideWhenUsed: true,
	},
	{
		type: 'table',
		id: 'TableNormal',
		name: 'Normal Table',
		isDefault: true,
		uiPriority: 99,
		semiHidden: true,
		unhideWhenUsed: true,
		tableXml:
			'<w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr>',
	},
	{
		type: 'numbering',
		id: 'NoList',
		name: 'No List',
		isDefault: true,
		uiPriority: 99,
		semiHidden: true,
		unhideWhenUsed: true,
	},
];

export function stylesXml(sheet: StyleSheet): string {
	return [
		'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
		`<w:styles xmlns:w="${W_NS}">`,
		`<w:docDefaults><w:rPrDefault><w:rPr>${runLook(sheet.run)}</w:rPr></w:rPrDefault>`,
		`<w:pPrDefault><w:pPr>${paragraphLook(sheet.paragraph)}</w:pPr></w:pPrDefault></w:docDefaults>`,
		...sheet.styles.map(styleXml),
		'</w:styles>',
	].join('');
}

/** A paragraph style and the character style linked to it (Word pairs them for partial use). */
export function linked(style: StyleSpec): StyleSpec[] {
	const charId = `${style.id}Char`;
	return [
		{ ...style, link: charId },
		{
			type: 'character',
			id: charId,
			name: `${style.name} Char`,
			basedOn: 'DefaultParagraphFont',
			link: style.id,
			uiPriority: style.uiPriority ?? 99,
			...(style.custom && { custom: true }),
			...(style.run && { run: style.run }),
		},
	];
}

/** Cell borders, shading and text of one region of a table style (`w:tblStylePr`). */
export function tableRegion(
	type: 'firstRow' | 'lastRow' | 'firstCol' | 'lastCol' | 'band1Horz' | 'band2Horz',
	look: {
		run?: RunLook;
		fill?: string;
		themeFill?: string;
		top?: Border;
		bottom?: Border;
		paragraph?: ParagraphLook;
	},
): string {
	const borders =
		look.top || look.bottom
			? `<w:tcBorders>${border('top', look.top)}${border('bottom', look.bottom)}</w:tcBorders>`
			: '';
	const fill = shading(look.fill, look.themeFill);
	return [
		`<w:tblStylePr w:type="${type}">`,
		look.paragraph ? `<w:pPr>${paragraphLook(look.paragraph)}</w:pPr>` : '',
		look.run ? `<w:rPr>${runLook(look.run)}</w:rPr>` : '',
		borders || fill ? `<w:tcPr>${borders}${fill}</w:tcPr>` : '',
		'</w:tblStylePr>',
	].join('');
}

export { border as borderXml };

/** A single line of `size` eighths of a point, `space` points from the text. */
export const rule = (color: string, size = 6, space = 0): Border => ({ size, space, color });

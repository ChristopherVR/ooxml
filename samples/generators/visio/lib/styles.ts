/**
 * The document part: fonts, the five style sheets every Visio drawing has and the sample's own
 * named styles. Every colour, weight and font a drawing uses lives in a style sheet here, so a
 * shape or master only names its style.
 */
import { FONTS } from '../../lib/fiction';
import { rootStyle, styleElement } from './root-style';
import { isThemeColor, paint } from './theme';
import { REL_NS, VISIO_NS, XML_HEADER, attrs, cell, colorCell, pt, rgb, row, section } from './xml';

export interface LineStyle {
	color?: string;
	/** Points. */
	weight?: number;
	/** 0 none, 1 solid, 2 dashed, 3 dotted, 4 dash-dot. */
	pattern?: number;
	/** Corner radius in inches. */
	rounding?: number;
	beginArrow?: number;
	endArrow?: number;
	arrowSize?: number;
}

export interface FillStyle {
	color?: string;
	/** 0 none, 1 solid. */
	pattern?: number;
	/** 0 to 1. */
	transparency?: number;
}

export interface TextStyle {
	font?: string;
	/** Points. */
	size?: number;
	color?: string;
	bold?: boolean;
	italic?: boolean;
	/** 0 left, 1 centre, 2 right. */
	align?: 0 | 1 | 2;
	/** 0 top, 1 middle, 2 bottom. */
	valign?: 0 | 1 | 2;
	/** Text block margins in points. */
	margin?: number;
	/** Opaque colour behind the text, for labels that sit on a line. */
	background?: string;
	/** 1 all caps. */
	caps?: boolean;
	/** Extra letter spacing in points. */
	tracking?: number;
	/** Line spacing as a multiple of the font size. */
	leading?: number;
}

export interface StyleDef {
	name: string;
	/** The style this one inherits from; Normal when absent. */
	based?: string;
	line?: LineStyle;
	fill?: FillStyle;
	text?: TextStyle;
}

const BUILT_IN = ['No Style', 'Text Only', 'None', 'Normal', 'Guide'] as const;

function lineCells(line: LineStyle | undefined): string[] {
	if (!line) return [];
	const cells: string[] = [];
	if (line.weight !== undefined) cells.push(cell('LineWeight', pt(line.weight), undefined, 'PT'));
	if (line.color) cells.push(colorCell('LineColor', line.color));
	if (line.pattern !== undefined) cells.push(cell('LinePattern', line.pattern));
	if (line.rounding !== undefined) cells.push(cell('Rounding', line.rounding));
	if (line.beginArrow !== undefined) cells.push(cell('BeginArrow', line.beginArrow));
	if (line.endArrow !== undefined) cells.push(cell('EndArrow', line.endArrow));
	if (line.arrowSize !== undefined)
		cells.push(cell('BeginArrowSize', line.arrowSize), cell('EndArrowSize', line.arrowSize));
	return cells;
}

function fillCells(fill: FillStyle | undefined): string[] {
	if (!fill) return [];
	const cells: string[] = [];
	if (fill.color) cells.push(colorCell('FillForegnd', fill.color));
	if (fill.pattern !== undefined) cells.push(cell('FillPattern', fill.pattern));
	if (fill.transparency !== undefined)
		cells.push(
			cell('FillForegndTrans', fill.transparency),
			cell('FillBkgndTrans', fill.transparency),
		);
	return cells;
}

function textBlockCells(text: TextStyle | undefined): string[] {
	if (!text) return [];
	const cells: string[] = [];
	if (text.margin !== undefined)
		for (const side of ['Left', 'Right', 'Top', 'Bottom'])
			cells.push(cell(`${side}Margin`, pt(text.margin), undefined, 'PT'));
	if (text.valign !== undefined) cells.push(cell('VerticalAlign', text.valign));
	// Visio stores a text background as the colour plus one; zero means none.
	if (text.background)
		cells.push(
			cell('TextBkgnd', text.background.toLowerCase(), `${rgb(text.background)}+1`),
			cell('TextBkgndTrans', 0),
		);
	return cells;
}

/** The Character row cells a text style sets; also used for the runs of mixed text. */
export function characterCells(text: TextStyle): string[] {
	const cells: string[] = [];
	if (text.font) cells.push(cell('Font', text.font));
	if (text.color) cells.push(colorCell('Color', text.color));
	if (text.bold !== undefined || text.italic !== undefined)
		cells.push(cell('Style', (text.bold ? 1 : 0) + (text.italic ? 2 : 0)));
	if (text.caps !== undefined) cells.push(cell('Case', text.caps ? 1 : 0));
	if (text.size !== undefined) cells.push(cell('Size', pt(text.size), undefined, 'PT'));
	if (text.tracking !== undefined)
		cells.push(cell('Letterspace', pt(text.tracking), undefined, 'PT'));
	return cells;
}

/** The Paragraph row cells a text style sets. */
export function paragraphCells(text: TextStyle): string[] {
	const cells: string[] = [];
	if (text.leading !== undefined) cells.push(cell('SpLine', -text.leading));
	if (text.align !== undefined) cells.push(cell('HorzAlign', text.align));
	return cells;
}

function textSections(text: TextStyle | undefined): string {
	if (!text) return '';
	const character = characterCells(text);
	const paragraph = paragraphCells(text);
	return (
		section('Character', character.length ? [row({ IX: 0 }, character)] : []) +
		section('Paragraph', paragraph.length ? [row({ IX: 0 }, paragraph)] : [])
	);
}

const enabled = [
	cell('EnableLineProps', 1),
	cell('EnableFillProps', 1),
	cell('EnableTextProps', 1),
	cell('HideForApply', 0),
];

/** The named styles of one drawing and the document part that declares them. */
export class StyleBook {
	private readonly ids = new Map<string, number>();

	constructor(
		private readonly defs: readonly StyleDef[],
		private readonly ink: string,
		/** The scheme cells that select the drawing's theme, when it has one. */
		private readonly themeSelection = '',
	) {
		BUILT_IN.forEach((name, index) => this.ids.set(name, index));
		defs.forEach((def, index) => {
			if (this.ids.has(def.name)) throw new Error(`Style ${def.name} is defined twice.`);
			this.ids.set(def.name, BUILT_IN.length + index);
		});
	}

	id(name: string): number {
		const id = this.ids.get(name);
		if (id === undefined) throw new Error(`Unknown style: ${name}`);
		return id;
	}

	/** A style's own value for one colour, or the one it inherits. */
	private inherited(name: string, pick: (def: StyleDef) => string | undefined): string | undefined {
		for (let def = this.defs.find((entry) => entry.name === name); def;) {
			const value = pick(def);
			if (value) return value;
			const parent = def.based;
			def = this.defs.find((entry) => entry.name === parent);
		}
		return undefined;
	}

	/**
	 * Local cells that make a page shape follow the theme: where its style's line or fill colour
	 * is one of the theme's accents, the cell is saved as the theme formula with the colour as
	 * its cached value. Visio resolves a style sheet's own formulas without the page's theme, so
	 * the formulas sit on the shapes, as they do in drawings Visio saves.
	 */
	themeCells(style: string | { line: string; fill: string }): string[] {
		const named = typeof style === 'string' ? { line: style, fill: style } : style;
		const cells: string[] = [];
		const line = this.inherited(named.line, (def) => def.line?.color);
		const fill = this.inherited(named.fill, (def) => def.fill?.color);
		const visible = (kind: 'line' | 'fill', from: string) =>
			this.inherited(from, (def) =>
				def[kind]?.pattern === undefined ? undefined : String(def[kind]?.pattern),
			) !== '0';
		if (line && isThemeColor(line) && visible('line', named.line))
			cells.push(paint('LineColor', line));
		if (fill && isThemeColor(fill) && visible('fill', named.fill))
			cells.push(paint('FillForegnd', fill));
		return cells;
	}

	private sheets(): string {
		const builtIn = [
			rootStyle(this.ink, this.themeSelection),
			styleElement(1, 'Text Only', 3, [
				...enabled,
				cell('LinePattern', 0),
				cell('FillPattern', 0),
				cell('VerticalAlign', 0),
				section('Paragraph', [row({ IX: 0 }, [cell('HorzAlign', 0)])]),
			]),
			styleElement(2, 'None', 3, [...enabled, cell('LinePattern', 0), cell('FillPattern', 0)]),
			styleElement(3, 'Normal', 0, [...enabled, ...textBlockCells({ margin: 4 })]),
			styleElement(4, 'Guide', 3, [
				...enabled,
				colorCell('LineColor', '#2F80C2'),
				cell('LinePattern', 23),
				cell('FillPattern', 0),
				cell('NonPrinting', 1),
			]),
		];
		const custom = this.defs.map((def) =>
			styleElement(this.id(def.name), def.name, this.id(def.based ?? 'Normal'), [
				...enabled,
				...lineCells(def.line),
				...fillCells(def.fill),
				...textBlockCells(def.text),
				textSections(def.text),
			]),
		);
		return `<StyleSheets>${[...builtIn, ...custom].join('')}</StyleSheets>`;
	}

	/** `visio/document.xml`. */
	documentXml(): string {
		const settings =
			'<DocumentSettings TopPage="0" DefaultTextStyle="3" DefaultLineStyle="3" DefaultFillStyle="3" DefaultGuideStyle="4">' +
			'<GlueSettings>9</GlueSettings><SnapSettings>65847</SnapSettings><SnapExtensions>34</SnapExtensions>' +
			'<SnapAngles/><DynamicGridEnabled>1</DynamicGridEnabled><ProtectStyles>0</ProtectStyles>' +
			'<ProtectShapes>0</ProtectShapes><ProtectMasters>0</ProtectMasters><ProtectBkgnds>0</ProtectBkgnds>' +
			'</DocumentSettings>';
		const faces = [FONTS.body, FONTS.display]
			.map((name) => `<FaceName${attrs({ NameU: name })}/>`)
			.join('');
		const sheet =
			'<DocumentSheet NameU="TheDoc" IsCustomNameU="1" Name="TheDoc" IsCustomName="1" LineStyle="0" FillStyle="0" TextStyle="0">' +
			cell('OutputFormat', 0) +
			cell('LockPreview', 0) +
			cell('PreviewQuality', 0) +
			cell('PreviewScope', 0) +
			cell('DocLangID', 'en-US') +
			'</DocumentSheet>';
		return (
			`${XML_HEADER}<VisioDocument xmlns="${VISIO_NS}" xmlns:r="${REL_NS}" xml:space="preserve">` +
			`${settings}<FaceNames>${faces}</FaceNames>${this.sheets()}${sheet}</VisioDocument>`
		);
	}
}

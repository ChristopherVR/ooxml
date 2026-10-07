/**
 * Shared parser for a paragraph-level default run-property text style
 * (`.../a:p/a:pPr/a:defRPr`): font size, bold, italic, Latin and East Asian
 * typefaces, and colour.
 *
 * ChartML repeats this exact shape in several places: a legend entry's
 * per-series text override (`c:legendEntry/c:txPr`), a data table's cell text
 * defaults (`c:dTable/c:txPr`), the chart-wide `c:chartSpace/c:txPr` and the
 * runs of a rich-text data label. They all read the same attributes, so the
 * resolution lives here once instead of being hand-rolled at each call site.
 *
 * @module utils/chart-def-rpr-style
 */
import type { PptxChartLegendTextStyle, XmlObject } from '../types';
import type { ResolveChartColor } from './chart-color-choice';
import { writeChartColorChoice } from './chart-color-choice';

interface XmlLookupLike {
	getChildByLocalName: (parent: XmlObject | undefined, name: string) => XmlObject | undefined;
}

interface ColorParserLike {
	parseColor: (fillNode: XmlObject | undefined, placeholderColor?: string) => string | undefined;
}

/** Resolve `<c:txPr>/a:p/a:pPr/a:defRPr` beneath a ChartML text-properties node. */
export function resolveTxPrDefRPr(
	txPr: XmlObject | undefined,
	xmlLookup: XmlLookupLike,
): XmlObject | undefined {
	const paragraph = xmlLookup.getChildByLocalName(txPr, 'p');
	const pPr = xmlLookup.getChildByLocalName(paragraph, 'pPr');
	return xmlLookup.getChildByLocalName(pPr, 'defRPr');
}

/**
 * Parse an already-resolved `a:defRPr` node's size/bold/italic/font/colour
 * into a flat text style. Returns `undefined` when the node is absent or
 * carries none of the recognised attributes.
 *
 * `resolveTypeface`, when provided, resolves a theme-font placeholder token
 * (`+mn-lt`, `+mj-lt`, `+mn-ea`, ...) to the deck's concrete theme face,
 * mirroring the slide-text path's `resolveThemeTypeface`. Without it, a
 * `+mn-lt` typeface parses through literally, which is not a usable CSS font
 * name.
 */
export function parseDefRPrTextStyle(
	defRPr: XmlObject | undefined,
	xmlLookup: XmlLookupLike,
	colorParser: ColorParserLike,
	resolveTypeface?: (raw: string) => string,
): PptxChartLegendTextStyle | undefined {
	if (!defRPr) {
		return undefined;
	}
	const style: PptxChartLegendTextStyle = {};

	const size = Number.parseInt(String(defRPr['@_sz'] ?? ''), 10);
	if (Number.isFinite(size)) {
		style.fontSize = size / 100;
	}
	if (defRPr['@_b'] !== undefined) {
		style.bold = defRPr['@_b'] === '1' || defRPr['@_b'] === 'true';
	}
	if (defRPr['@_i'] !== undefined) {
		style.italic = defRPr['@_i'] === '1' || defRPr['@_i'] === 'true';
	}
	const latin = xmlLookup.getChildByLocalName(defRPr, 'latin');
	if (latin?.['@_typeface']) {
		const raw = String(latin['@_typeface']);
		style.fontFamily = resolveTypeface ? resolveTypeface(raw) : raw;
	}
	const ea = xmlLookup.getChildByLocalName(defRPr, 'ea');
	if (ea?.['@_typeface']) {
		const raw = String(ea['@_typeface']);
		style.eastAsiaFontFamily = resolveTypeface ? resolveTypeface(raw) : raw;
	}
	const color = colorParser.parseColor(xmlLookup.getChildByLocalName(defRPr, 'solidFill'));
	if (color) {
		style.color = color;
	}

	return Object.keys(style).length > 0 ? style : undefined;
}

/**
 * The chart-wide text style: `c:chartSpace/c:txPr` over the theme's minor
 * fonts (`+mn-lt`, `+mn-ea`), which is what PowerPoint draws chart text with
 * when nothing closer sets a face. A token the theme does not define is left
 * out rather than passed on as a font name.
 */
export function parseChartWideTextStyle(
	chartSpace: XmlObject | undefined,
	xmlLookup: XmlLookupLike,
	colorParser: ColorParserLike,
	resolveTypeface: (raw: string) => string,
): PptxChartLegendTextStyle | undefined {
	const themeFace = (token: string) => {
		const face = resolveTypeface(token);
		return face && !face.startsWith('+') ? face : undefined;
	};
	const latin = themeFace('+mn-lt');
	const eastAsia = themeFace('+mn-ea');
	const style: PptxChartLegendTextStyle = {
		...(latin ? { fontFamily: latin } : {}),
		...(eastAsia ? { eastAsiaFontFamily: eastAsia } : {}),
		...parseDefRPrTextStyle(
			resolveTxPrDefRPr(xmlLookup.getChildByLocalName(chartSpace, 'txPr'), xmlLookup),
			xmlLookup,
			colorParser,
			resolveTypeface,
		),
	};
	return Object.keys(style).length > 0 ? style : undefined;
}

/**
 * The font a rich-text label (`c:tx/c:rich`) is drawn with: its first run's
 * (or field's) `a:rPr` over the paragraph's `a:pPr/a:defRPr`. Only the first
 * run counts, since the label is drawn as one piece of text.
 */
export function parseRichTextStyle(
	rich: XmlObject | undefined,
	xmlLookup: XmlLookupLike,
	colorParser: ColorParserLike,
	resolveTypeface?: (raw: string) => string,
): PptxChartLegendTextStyle | undefined {
	const paragraph = xmlLookup.getChildByLocalName(rich, 'p');
	const defRPr = xmlLookup.getChildByLocalName(
		xmlLookup.getChildByLocalName(paragraph, 'pPr'),
		'defRPr',
	);
	// A label written only as a field (`a:fld`, such as [VALUE]) carries its font there.
	const run =
		xmlLookup.getChildByLocalName(paragraph, 'r') ??
		xmlLookup.getChildByLocalName(paragraph, 'fld');
	const rPr = xmlLookup.getChildByLocalName(run, 'rPr');
	const style = {
		...parseDefRPrTextStyle(defRPr, xmlLookup, colorParser, resolveTypeface),
		...parseDefRPrTextStyle(rPr, xmlLookup, colorParser, resolveTypeface),
	};
	return Object.keys(style).length > 0 ? style : undefined;
}

/**
 * Serialize a flat text style back into a `c:txPr` node (CT_TextBody):
 * `a:bodyPr` and `a:lstStyle` as empty placeholders, and the parsed attributes
 * inside `a:p/a:pPr/a:defRPr`. Shared by the legend-entry and data-table
 * writers, the two ChartML locations that use this exact shape.
 *
 * Returns `undefined` when `style` is absent or empty, meaning: leave any
 * authored `c:txPr` untouched.
 */
export function buildDefRPrTextProperties(
	style: PptxChartLegendTextStyle | undefined,
	authoredDefRPr: XmlObject | undefined,
	resolveColor?: ResolveChartColor,
): XmlObject | undefined {
	if (!style || Object.keys(style).length === 0) {
		return undefined;
	}
	const rPr: XmlObject = {};
	if (style.fontSize !== undefined) {
		rPr['@_sz'] = String(Math.round(style.fontSize * 100));
	}
	if (style.bold !== undefined) {
		rPr['@_b'] = style.bold ? '1' : '0';
	}
	if (style.italic !== undefined) {
		rPr['@_i'] = style.italic ? '1' : '0';
	}
	if (style.color) {
		rPr['a:solidFill'] = authoredDefRPr?.['a:solidFill'];
		writeChartColorChoice(rPr, 'a:solidFill', style.color, resolveColor);
	}
	if (style.fontFamily) {
		rPr['a:latin'] = { '@_typeface': style.fontFamily };
	}
	// The East Asian face is not editable, so an authored a:ea (often a theme
	// token such as +mn-ea, which parsing resolves) is written back unchanged.
	if (authoredDefRPr?.['a:ea']) {
		rPr['a:ea'] = authoredDefRPr['a:ea'];
	} else if (style.eastAsiaFontFamily) {
		rPr['a:ea'] = { '@_typeface': style.eastAsiaFontFamily };
	}
	return { 'a:bodyPr': {}, 'a:lstStyle': {}, 'a:p': { 'a:pPr': { 'a:defRPr': rPr } } };
}

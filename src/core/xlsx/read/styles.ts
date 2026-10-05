import { parseXml, type XmlElement } from '../../xml/index.js';
import type { CellStyle, DifferentialStyle, Workbook } from '../model.js';
import { styleKey } from '../styles.js';
import { defaultCellStyle } from '../workbook.js';
import { builtinFormat } from './builtin-formats.js';
import {
	parseAlignment,
	parseBorder,
	parseFill,
	parseFont,
	parseIndexedPalette,
	parseProtection,
} from './style-parts.js';
import { att, boolAttr, numAttr, xChildren, xFirst } from './xml-util.js';

export interface ParsedStyles {
	styles: CellStyle[];
	/** `cellXfs` index to {@link Workbook.styles} index. */
	xfMap: number[];
	namedStyles: Workbook['namedStyles'];
	/** Differential formats by `dxfId`. */
	dxfs: DifferentialStyle[];
	/** The legacy palette override, if the part has one. */
	palette?: string[];
}

/** Reads a `dxf` (conditional formatting, tables) into a differential style. */
export function parseDxf(
	dxf: XmlElement,
	numFmts: ReadonlyMap<number, string>,
	palette: readonly string[] | undefined,
): DifferentialStyle {
	const style: DifferentialStyle = {};
	const font = xFirst(dxf, 'font');
	if (font) style.font = parseFont(font, palette, true);
	const fill = xFirst(dxf, 'fill');
	if (fill) style.fill = parseFill(fill, palette, true);
	const border = xFirst(dxf, 'border');
	if (border) style.border = parseBorder(border, palette);
	const numFmt = xFirst(dxf, 'numFmt');
	if (numFmt) {
		const code = att(numFmt, 'formatCode') ?? numFmts.get(numAttr(numFmt, 'numFmtId') ?? -1);
		if (code) style.numFmt = code;
	}
	return style;
}

/** Reads `styles.xml`; an absent part gives the default style table. */
export function parseStyles(xml: string | undefined, warnings: string[]): ParsedStyles {
	const fallback = defaultCellStyle();
	if (!xml)
		return {
			styles: [fallback],
			xfMap: [0],
			namedStyles: [{ name: 'Normal', style: defaultCellStyle(), builtinId: 0 }],
			dxfs: [],
		};
	const root = parseXml(xml, { label: 'XLSX styles' }).documentElement;
	const palette = parseIndexedPalette(root);
	const numFmts = new Map<number, string>();
	for (const node of xChildren(xFirst(root, 'numFmts') ?? root, 'numFmt')) {
		const id = numAttr(node, 'numFmtId');
		const code = att(node, 'formatCode');
		if (id !== undefined && code !== undefined) numFmts.set(id, code);
	}
	const formatOf = (id: number | undefined): string => {
		if (id === undefined) return 'General';
		const code = numFmts.get(id) ?? builtinFormat(id);
		if (code === undefined) {
			warnings.push(`Unknown number format id ${id}; using General.`);
			return 'General';
		}
		return code;
	};
	const fonts = xChildren(xFirst(root, 'fonts') ?? root, 'font').map((node) =>
		parseFont(node, palette),
	);
	const fills = xChildren(xFirst(root, 'fills') ?? root, 'fill').map((node) =>
		parseFill(node, palette),
	);
	const borders = xChildren(xFirst(root, 'borders') ?? root, 'border').map((node) =>
		parseBorder(node, palette),
	);
	const resolve = (xf: XmlElement): CellStyle => {
		const style: CellStyle = {
			font: fonts[numAttr(xf, 'fontId') ?? 0] ?? fonts[0] ?? fallback.font,
			fill: fills[numAttr(xf, 'fillId') ?? 0] ?? { type: 'pattern', pattern: 'none' },
			border: borders[numAttr(xf, 'borderId') ?? 0] ?? {},
			numFmt: formatOf(numAttr(xf, 'numFmtId')),
		};
		const alignment = parseAlignment(xFirst(xf, 'alignment'));
		if (alignment) style.alignment = alignment;
		const protection = parseProtection(xFirst(xf, 'protection'));
		if (protection) style.protection = protection;
		if (boolAttr(xf, 'quotePrefix')) style.quotePrefix = true;
		if (boolAttr(xf, 'pivotButton')) style.pivotButton = true;
		return style;
	};

	const styleXfs = xChildren(xFirst(root, 'cellStyleXfs') ?? root, 'xf').map(resolve);
	const namesByXf = new Map<number, string>();
	const namedStyles: Workbook['namedStyles'] = [];
	for (const node of xChildren(xFirst(root, 'cellStyles') ?? root, 'cellStyle')) {
		const xfId = numAttr(node, 'xfId') ?? 0;
		const name = att(node, 'name') ?? `Style ${xfId}`;
		const style = styleXfs[xfId] ?? fallback;
		const entry: Workbook['namedStyles'][number] = { name, style: structuredClone(style) };
		const builtinId = numAttr(node, 'builtinId');
		if (builtinId !== undefined) entry.builtinId = builtinId;
		if (!namesByXf.has(xfId)) namesByXf.set(xfId, name);
		namedStyles.push(entry);
	}
	const normalAt = namedStyles.findIndex((entry) => entry.builtinId === 0);
	if (normalAt > 0) namedStyles.unshift(...namedStyles.splice(normalAt, 1));
	if (!namedStyles.some((entry) => entry.name === 'Normal'))
		namedStyles.unshift({ name: 'Normal', style: styleXfs[0] ?? defaultCellStyle(), builtinId: 0 });

	const styles: CellStyle[] = [];
	const keys = new Map<string, number>();
	const xfMap: number[] = [];
	for (const xf of xChildren(xFirst(root, 'cellXfs') ?? root, 'xf')) {
		const style = resolve(xf);
		const xfId = numAttr(xf, 'xfId') ?? 0;
		const named = namesByXf.get(xfId);
		if (named && xfId !== 0 && named !== 'Normal') style.cellStyleName = named;
		const key = styleKey(style);
		let index = keys.get(key);
		if (index === undefined) {
			index = styles.length;
			styles.push(style);
			keys.set(key, index);
		}
		xfMap.push(index);
	}
	if (!styles.length) {
		styles.push(fallback);
		xfMap.push(0);
	}
	const dxfs = xChildren(xFirst(root, 'dxfs') ?? root, 'dxf').map((node) =>
		parseDxf(node, numFmts, palette),
	);
	const result: ParsedStyles = { styles, xfMap, namedStyles, dxfs };
	if (palette) result.palette = palette;
	return result;
}

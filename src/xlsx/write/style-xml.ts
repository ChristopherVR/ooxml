import type { Alignment, Border, Color, Fill, Font, Protection } from '../model.js';
import { attrs, el, escapeAttr } from './xml-out.js';

/** `RRGGBB` becomes `FFRRGGBB`; SpreadsheetML colours are ARGB. */
const argb = (rgb: string): string => (rgb.length === 6 ? `FF${rgb}` : rgb).toUpperCase();

export function colorXml(color: Color | undefined, name = 'color'): string {
	if (!color) return '';
	if (color.rgb) return el(name, { rgb: argb(color.rgb), tint: color.tint });
	if (color.theme !== undefined) return el(name, { theme: color.theme, tint: color.tint });
	if (color.indexed !== undefined) return el(name, { indexed: color.indexed, tint: color.tint });
	if (color.auto) return el(name, { auto: true });
	return '';
}

/**
 * A font: `<font>` children for cell formats, `<rPr>` with `rFont` for rich text runs. In a
 * differential font an explicit `false` is written as `val="0"`.
 */
export function fontXml(font: Font, tag: 'font' | 'rPr' = 'font', differential = false): string {
	const flag = (name: string, value: boolean | undefined) =>
		value === true ? `<${name}/>` : value === false && differential ? `<${name} val="0"/>` : '';
	let out = flag('b', font.bold) + flag('i', font.italic) + flag('strike', font.strike);
	if (font.underline)
		out += font.underline === 'single' ? '<u/>' : el('u', { val: font.underline });
	if (font.vertAlign) out += el('vertAlign', { val: font.vertAlign });
	if (font.size !== undefined) out += el('sz', { val: font.size });
	out += colorXml(font.color);
	if (font.name) out += el(tag === 'rPr' ? 'rFont' : 'name', { val: font.name });
	if (font.family !== undefined) out += el('family', { val: font.family });
	if (font.scheme) out += el('scheme', { val: font.scheme });
	return `<${tag}>${out}</${tag}>`;
}

export function fillXml(fill: Fill, differential = false): string {
	if (fill.type === 'gradient') {
		const stops = fill.stops
			.map((stop) => `<stop position="${stop.position}">${colorXml(stop.color)}</stop>`)
			.join('');
		return `<fill>${el('gradientFill', { type: fill.gradient === 'path' ? 'path' : undefined, degree: fill.degree }, stops)}</fill>`;
	}
	if (differential && fill.pattern === 'solid') {
		const color = fill.fgColor ?? fill.bgColor;
		return `<fill><patternFill patternType="solid">${colorXml(color, 'fgColor')}${colorXml(color, 'bgColor')}</patternFill></fill>`;
	}
	const colors = colorXml(fill.fgColor, 'fgColor') + colorXml(fill.bgColor, 'bgColor');
	return `<fill>${el('patternFill', { patternType: fill.pattern }, colors)}</fill>`;
}

export function borderXml(border: Border): string {
	const edge = (name: string, value: Border['left']) =>
		value ? el(name, { style: value.style }, colorXml(value.color)) : `<${name}/>`;
	const inner =
		edge('left', border.left) +
		edge('right', border.right) +
		edge('top', border.top) +
		edge('bottom', border.bottom) +
		edge('diagonal', border.diagonal);
	return `<border${attrs({ diagonalUp: border.diagonalUp || undefined, diagonalDown: border.diagonalDown || undefined })}>${inner}</border>`;
}

export function alignmentXml(alignment: Alignment | undefined): string {
	if (!alignment || !Object.keys(alignment).length) return '';
	return el('alignment', {
		horizontal: alignment.horizontal,
		vertical: alignment.vertical,
		textRotation: alignment.textRotation,
		wrapText: alignment.wrapText || undefined,
		indent: alignment.indent,
		shrinkToFit: alignment.shrinkToFit || undefined,
		readingOrder: alignment.readingOrder,
	});
}

export function protectionXml(protection: Protection | undefined): string {
	if (!protection || !Object.keys(protection).length) return '';
	return el('protection', { locked: protection.locked, hidden: protection.hidden });
}

export const numFmtXml = (id: number, code: string): string =>
	`<numFmt numFmtId="${id}" formatCode="${escapeAttr(code)}"/>`;

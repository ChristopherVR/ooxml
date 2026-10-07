// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
import { elements, first, getW, type XmlElement, WORD_NS, WORD_DATE_UTC_NS } from './xml';
import { isWordHighlightToken } from './highlight';
import { isValidLanguageTag } from './language';
import { isWordUnderlineStyle } from './underline';
import { isThemeColorToken } from './theme-color';
import { isLigatures, WORD_2010_NS } from './ligatures';
import { isStVerticalAlignRun } from './generated/wml-simple-types';
import {
	parseHalfPoints,
	parseHexColor,
	parseSignedTwips,
	parseSignedHalfPoints,
	parseTextScale,
} from './simple-types';

const modeledRunProperties = new Set([
	'b',
	'i',
	'bCs',
	'iCs',
	'szCs',
	'strike',
	'u',
	'highlight',
	'vertAlign',
	'sz',
	'rFonts',
	'color',
	'lang',
	'rtl',
	'rStyle',
	'caps',
	'smallCaps',
	'dstrike',
	'vanish',
	'spacing',
	'w',
	'kern',
	'position',
	'shd',
]);
const simpleToggleProperties = [
	'b',
	'i',
	'bCs',
	'iCs',
	'strike',
	'caps',
	'smallCaps',
	'dstrike',
	'vanish',
];
const allowedRunPropertyAttributes: Record<string, string[]> = {
	rFonts: [
		'ascii',
		'hAnsi',
		'eastAsia',
		'cs',
		'asciiTheme',
		'hAnsiTheme',
		'eastAsiaTheme',
		'cstheme',
	],
	lang: ['val', 'eastAsia', 'bidi'],
	highlight: ['val'],
	vertAlign: ['val'],
	sz: ['val'],
	szCs: ['val'],
	rtl: ['val'],
	rStyle: ['val'],
	spacing: ['val'],
	w: ['val'],
	kern: ['val'],
	position: ['val'],
	u: ['val', 'color'],
	color: ['val', 'themeColor', 'themeTint', 'themeShade'],
	shd: ['val', 'fill', 'color', 'themeFill', 'themeFillTint', 'themeFillShade'],
};
function hasUnexpectedAttributes(
	element: XmlElement,
	allowed: string[],
	allowDateUtc = false,
): boolean {
	for (const attribute of Array.from(element.attributes)) {
		if (attribute.namespaceURI === 'http://www.w3.org/2000/xmlns/') continue;
		if (
			allowDateUtc &&
			attribute.namespaceURI === WORD_DATE_UTC_NS &&
			attribute.localName === 'dateUtc'
		)
			continue;
		if (attribute.namespaceURI !== WORD_NS || !allowed.includes(attribute.localName)) return true;
	}
	return false;
}

/** Whether a run's `rPr` contains any property/attribute/value this writer does not model. */
export function runHasUnknownProperties(run: XmlElement): boolean {
	return runPropertiesHaveUnknownContent(first(run, 'rPr'));
}

export function runPropertiesHaveUnknownContent(properties: XmlElement | undefined): boolean {
	if (!properties) return false;
	if (hasUnexpectedAttributes(properties, [])) return true;
	for (const node of Array.from(properties.childNodes)) {
		if (node.nodeType !== 1) {
			if (node.nodeType === 3 && node.textContent?.trim()) return true;
			continue;
		}
		const property = node as XmlElement;
		if (property.namespaceURI === WORD_NS && property.localName === 'rPrChange') {
			// Its complete prior subtree is modeled as XML, including unsupported historical properties.
			if (
				hasUnexpectedAttributes(property, ['id', 'author', 'date'], true) ||
				Array.from(property.childNodes).some(
					(child) => child.nodeType === 3 && child.textContent?.trim(),
				) ||
				elements(property).length !== 1 ||
				!first(property, 'rPr') ||
				!getW(property, 'id') ||
				!getW(property, 'author')
			)
				return true;
			continue;
		}
		if (property.namespaceURI === WORD_2010_NS && property.localName === 'ligatures') {
			if (!isLigatures(property.getAttributeNS(WORD_2010_NS, 'val')) || elements(property).length)
				return true;
			if (
				Array.from(property.attributes).some(
					(attribute) =>
						attribute.namespaceURI !== 'http://www.w3.org/2000/xmlns/' &&
						(attribute.namespaceURI !== WORD_2010_NS || attribute.localName !== 'val'),
				)
			)
				return true;
			continue;
		}
		if (property.namespaceURI !== WORD_NS || !modeledRunProperties.has(property.localName))
			return true;
		const allowed = simpleToggleProperties.includes(property.localName)
			? ['val']
			: (allowedRunPropertyAttributes[property.localName] ?? []);
		if (hasUnexpectedAttributes(property, allowed) || elements(property).length > 0) return true;
		const value = getW(property, 'val');
		if (property.localName === 'w' && parseTextScale(value ?? '100') === undefined) return true;
		if (property.localName === 'kern' && parseHalfPoints(value) === undefined) return true;
		if (property.localName === 'position' && parseSignedHalfPoints(value) === undefined)
			return true;
		if (property.localName === 'highlight' && value && !isWordHighlightToken(value)) return true;
		if (
			property.localName === 'rtl' &&
			value &&
			!['1', 'true', 'on', '0', 'false', 'off', 'no'].includes(value)
		)
			return true;
		if (
			property.localName === 'lang' &&
			['val', 'eastAsia', 'bidi'].some((key) => {
				const language = getW(property, key);
				return language !== undefined && !isValidLanguageTag(language);
			})
		)
			return true;
		if (property.localName === 'vertAlign' && !isStVerticalAlignRun(value)) return true;
		if (
			property.localName === 'u' &&
			value &&
			!isWordUnderlineStyle(value) &&
			!['0', 'false', 'off', '1', 'true', 'on'].includes(value)
		)
			return true;
		if (property.localName === 'u') {
			const color = getW(property, 'color');
			if (color && parseHexColor(color) === undefined) return true;
		}
		if (
			['sz', 'szCs'].includes(property.localName) &&
			value &&
			parseHalfPoints(value) === undefined
		)
			return true;
		if (property.localName === 'spacing' && value && parseSignedTwips(value) === undefined)
			return true;
		if (property.localName === 'color' || property.localName === 'shd') {
			const fillValue = property.localName === 'shd' ? getW(property, 'fill') : value;
			if (fillValue && parseHexColor(fillValue) === undefined) return true;
			const themeValue = getW(property, property.localName === 'shd' ? 'themeFill' : 'themeColor');
			if (themeValue && !isThemeColorToken(themeValue)) return true;
		}
	}
	return false;
}

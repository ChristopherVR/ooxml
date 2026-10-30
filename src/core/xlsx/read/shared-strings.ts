import { LiteDeclinedError, NS, named, type XmlElement } from '../../xml/index';
import type { PhoneticProperties, PhoneticRun, PhoneticText, RichTextRun } from '../model';
import { checkDataXml, parseDataXml } from './limits';
import { rootChildren } from './sheet-data-source';
import { parseFont } from './style-parts';
import { att, numAttr, xChildren, xFirst, xText } from './xml-util';

export interface SharedString {
	text: string;
	runs?: RichTextRun[];
	/** Phonetic guide runs (`rPh`) and `phoneticPr`, when the string has them. */
	phonetic?: PhoneticText;
}

const PHONETIC_TYPES = [
	'halfwidthKatakana',
	'fullwidthKatakana',
	'Hiragana',
	'noConversion',
] as const;
const PHONETIC_ALIGNMENTS = ['noControl', 'left', 'center', 'distributed'] as const;

/** The member of `allowed` that `value` spells, or undefined (unknown values are dropped). */
const oneOf = <T extends string>(allowed: readonly T[], value: string | undefined) =>
	allowed.find((item) => item === value);

/** Reads the phonetic guide of a `CT_Rst`: its `rPh` runs and `phoneticPr` settings. */
function parsePhonetic(element: XmlElement): PhoneticText | undefined {
	const runs: PhoneticRun[] = xChildren(element, 'rPh').map((run) => {
		const sb = numAttr(run, 'sb') ?? 0;
		return { text: xText(xFirst(run, 't')), sb, eb: numAttr(run, 'eb') ?? sb };
	});
	const pr = xFirst(element, 'phoneticPr');
	if (!runs.length && !pr) return undefined;
	const phonetic: PhoneticText = { runs };
	if (pr) {
		const properties: PhoneticProperties = {};
		const fontId = numAttr(pr, 'fontId');
		if (fontId !== undefined) properties.fontId = fontId;
		const type = oneOf(PHONETIC_TYPES, att(pr, 'type'));
		if (type) properties.type = type;
		const alignment = oneOf(PHONETIC_ALIGNMENTS, att(pr, 'alignment'));
		if (alignment) properties.alignment = alignment;
		if (Object.keys(properties).length) phonetic.properties = properties;
	}
	return phonetic;
}

/**
 * Reads a `CT_Rst` (`<si>` or `<is>`): plain `<t>` or rich `<r>` runs, and the phonetic guide
 * (`rPh` runs and `phoneticPr`), kept on `phonetic`.
 */
export function parseRichString(
	element: XmlElement,
	palette: readonly string[] | undefined,
): SharedString {
	const phonetic = parsePhonetic(element);
	const withPhonetic = (value: SharedString): SharedString =>
		phonetic ? { ...value, phonetic } : value;
	const runs = xChildren(element, 'r');
	if (!runs.length) return withPhonetic({ text: xText(xFirst(element, 't')) });
	const parsed: RichTextRun[] = runs.map((run) => {
		const out: RichTextRun = { text: xText(xFirst(run, 't')) };
		const rPr = xFirst(run, 'rPr');
		if (rPr) {
			const font = parseFont(rPr, palette);
			if (Object.keys(font).length) out.font = font;
		}
		return out;
	});
	const text = parsed.map((run) => run.text).join('');
	return withPhonetic({ text, runs: parsed });
}

const LABEL = 'XLSX shared strings';

/**
 * Reads the shared strings part. The strings are read one `<si>` at a time with the light reader
 * (a DOM of a large table cost about 2 KiB per string), after the element and nesting limits the
 * DOM parser checks; a part the light reader declines is parsed whole, with the same result.
 */
export function parseSharedStrings(
	xml: string | undefined,
	palette: readonly string[] | undefined,
	maxXmlElements?: number,
): SharedString[] {
	if (!xml) return [];
	const items = rootChildren(xml);
	if (items) {
		checkDataXml(xml, LABEL, maxXmlElements);
		try {
			const out: SharedString[] = [];
			for (const si of items) if (named(si, 'si', NS.x)) out.push(parseRichString(si, palette));
			return out;
		} catch (error) {
			if (!(error instanceof LiteDeclinedError)) throw error;
		}
	}
	const root = parseDataXml(xml, LABEL, maxXmlElements).documentElement;
	return xChildren(root, 'si').map((si) => parseRichString(si, palette));
}

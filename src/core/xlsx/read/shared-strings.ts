import { LiteDeclinedError, NS, named, type XmlElement } from '../../xml/index';
import type { RichTextRun } from '../model';
import { checkDataXml, parseDataXml } from './limits';
import { rootChildren } from './sheet-data-source';
import { parseFont } from './style-parts';
import { xChildren, xFirst, xText } from './xml-util';

export interface SharedString {
	text: string;
	runs?: RichTextRun[];
}

/**
 * Reads a `CT_Rst` (`<si>` or `<is>`): plain `<t>` or rich `<r>` runs. Phonetic runs (`rPh`)
 * and properties are ignored.
 */
export function parseRichString(
	element: XmlElement,
	palette: readonly string[] | undefined,
): SharedString {
	const runs = xChildren(element, 'r');
	if (!runs.length) return { text: xText(xFirst(element, 't')) };
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
	return { text, runs: parsed };
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

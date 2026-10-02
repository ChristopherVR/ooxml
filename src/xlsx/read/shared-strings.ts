import { parseXml, type XmlElement } from '../../xml/index.js';
import type { RichTextRun } from '../model.js';
import { parseFont } from './style-parts.js';
import { xChildren, xFirst, xText } from './xml-util.js';

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

export function parseSharedStrings(
	xml: string | undefined,
	palette: readonly string[] | undefined,
): SharedString[] {
	if (!xml) return [];
	const root = parseXml(xml, { label: 'XLSX shared strings' }).documentElement;
	return xChildren(root, 'si').map((si) => parseRichString(si, palette));
}

import type { Paragraph, TextRun } from './model';
import { parseDirectRunProperties } from './run-properties';
import { first, parseXml } from './xml';

/** A formatting-only run representing the saved paragraph mark. */
export function paragraphMarkRun(paragraph: Paragraph): TextRun {
	const source = paragraph.sourceParagraphPropertiesXml;
	const properties = source ? first(parseXml(source).documentElement, 'rPr') : undefined;
	return { text: '', ...parseDirectRunProperties(properties) };
}

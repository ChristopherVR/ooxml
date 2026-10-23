// Shared helpers of the `cx:chartSpace` writer: chartex elements and attributes, and the shape,
// text and number-format pieces written through the classic chart writers.
import type { ChartNumberFormat, ChartShapeProperties, ChartTextBody } from '../model-series';
import { shapePropertiesInnerXml, textBodyInnerXml } from '../write-shape';
import { chartNumber, escapeAttribute, escapeText, type ChartWriteContext } from '../write-util';
import type { ChartExFormula } from './model';

export { raw, type ChartWriteContext } from '../write-util';

/** An attribute value: booleans as `1`/`0`, numbers canonical, `auto` and text as written. */
export type AttributeValue = string | number | boolean | undefined;

/** ` name="value"` for every defined entry, in the order given (the schema's). */
export function attributes(entries: [string, AttributeValue][]): string {
	return entries
		.map(([name, value]) => {
			if (value === undefined) return '';
			const text =
				typeof value === 'boolean'
					? value
						? '1'
						: '0'
					: typeof value === 'number'
						? chartNumber(value)
						: escapeAttribute(value);
			return ` ${name}="${text}"`;
		})
		.join('');
}

/** `<cx:local attrs>content</cx:local>`, or the empty element when there is no content. */
export const cxElement = (local: string, content = '', attrs = ''): string =>
	content ? `<cx:${local}${attrs}>${content}</cx:${local}>` : `<cx:${local}${attrs}/>`;

/** `<cx:local>text</cx:local>` with the text escaped, for a defined value. */
export const cxText = (local: string, text: string | undefined, attrs = ''): string =>
	text === undefined ? '' : `<cx:${local}${attrs}>${escapeText(text)}</cx:${local}>`;

/** `<cx:local val="..."/>` for a defined value. */
export const cxVal = (local: string, value: AttributeValue): string =>
	value === undefined ? '' : `<cx:${local}${attributes([['val', value]])}/>`;

/** A formula element (`cx:f`, `cx:nf`). */
export const formulaXml = (local: string, formula: ChartExFormula | undefined): string =>
	formula ? cxText(local, formula.formula, attributes([['dir', formula.direction]])) : '';

/** `cx:spPr`. */
export const shapeXml = (context: ChartWriteContext, shape: ChartShapeProperties | undefined) =>
	shape ? cxElement('spPr', shapePropertiesInnerXml(context, shape)) : '';

/** `cx:txPr`, or `cx:rich` with `local`. */
export const textBodyXml = (
	context: ChartWriteContext,
	body: ChartTextBody | undefined,
	local = 'txPr',
): string => (body ? cxElement(local, textBodyInnerXml(context, body)) : '');

/** `cx:numFmt`. */
export function numberFormatXml(format: ChartNumberFormat | undefined): string {
	if (!format) return '';
	return cxElement(
		'numFmt',
		'',
		attributes([
			['formatCode', format.formatCode],
			['sourceLinked', format.sourceLinked],
		]),
	);
}

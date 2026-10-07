import { parseDrawingColorIn } from '../diagram/drawing-color';
import { parseDrawingFill, parseDrawingLine } from '../diagram/drawing-fill';
import { NS, first, type XmlElement } from '../xml/index';
import type { ChartStyleEntry } from './style-definition';

/** Common DrawingML properties in chart styles and directly formatted chart elements. */
export function readChartAppearance(
	text: XmlElement | undefined,
	properties: XmlElement | undefined,
): Omit<ChartStyleEntry, 'sourceXml'> {
	const out: Omit<ChartStyleEntry, 'sourceXml'> = {};
	const size = Number(text?.getAttribute('sz') ?? NaN);
	if (Number.isFinite(size) && size > 0) out.fontSize = size / 100;
	for (const [attribute, key] of [
		['b', 'bold'],
		['i', 'italic'],
	] as const) {
		const value = text?.getAttribute(attribute);
		if (value !== null && value !== undefined) out[key] = value === '1' || value === 'true';
	}
	const typeface = first(text, 'latin', NS.a)?.getAttribute('typeface');
	if (typeface) out.typeface = typeface;
	const color = parseDrawingColorIn(first(text, 'solidFill', NS.a));
	if (color) out.textColor = color;
	const fill = parseDrawingFill(properties);
	const line = parseDrawingLine(first(properties, 'ln', NS.a));
	if (fill) out.fill = fill;
	if (line) out.line = line;
	return out;
}

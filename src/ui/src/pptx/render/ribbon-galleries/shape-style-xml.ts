import { themeShapeStyle, type DrawingColor, type ShapeStyleColumn } from 'ooxml-core/drawingml';
import type { XmlObject } from 'ooxml-core/pptx';

function styleRef(idx: number | string, color: DrawingColor): XmlObject {
	const scheme: XmlObject = { '@_val': color.value };
	for (const { name, value } of color.transforms) scheme[`a:${name}`] = { '@_val': value };
	return { '@_idx': String(idx), 'a:schemeClr': scheme };
}

/** Adapts the shared theme references to the presentation model's object tree. */
export function themeShapeStyleXml(row: number, column: ShapeStyleColumn): XmlObject {
	const style = themeShapeStyle(row, column);
	if (!style) throw new RangeError('Unknown theme style row');
	return {
		'a:lnRef': styleRef(style.line.idx, style.line.color),
		'a:fillRef': styleRef(style.fill.idx, style.fill.color),
		'a:effectRef': styleRef(style.effect.idx, style.effect.color),
		'a:fontRef': styleRef(style.font.idx, style.font.color),
	};
}

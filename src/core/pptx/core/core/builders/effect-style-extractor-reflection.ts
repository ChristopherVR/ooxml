import type { ShapeStyle, XmlObject } from '../../types';
import { readObjectEffect, reflectionStyle } from './effect-style-from-model';

/** `a:reflection` fields of a shape or text run, read through the shared effect reader. */
export function extractReflectionAttributes(
	node: XmlObject,
	emuPerPx: number,
): Partial<ShapeStyle> {
	return reflectionStyle(readObjectEffect('reflection', node), node, emuPerPx);
}

/**
 * A shape's GEOMETRY FAMILY, and the matching rule PowerPoint applies to it.
 *
 * Measured on PowerPoint 16 (`CreateVideo`, 60 fps, 720p, Morph by object, a
 * slide duplicated so every shape keeps its `a16:creationId`):
 *
 *  - a rectangle turned into an ellipse, a rounded rectangle or a freeform
 *    (Edit Points), and an isosceles triangle turned into a pentagon, do NOT
 *    morph: the old shape fades out where it stood (gone by about 25% of the
 *    morph) and the new one fades in where it lands (from about 45%), exactly
 *    the unmatched pair's dissolve;
 *  - two freeforms whose nodes differ, even in number, still morph;
 *  - the same rectangle -> ellipse change on a `!!`-named pair morphs its
 *    outline (see `morph-outline-blend`).
 *
 * So identity alone pairs two shapes only when their geometry family (the
 * preset, or "custom" for any authored path) is the same; the `!!` name
 * overrides it.
 *
 * @module render/morph-geometry-key
 */
import type { PptxElement } from 'ooxml-core/pptx';
import { getShapeType, hasShapeProperties } from 'ooxml-core/pptx';

/** A shape's geometry family: its preset, or `custom` for an authored path. */
export function morphGeometryKey(element: PptxElement): string | undefined {
	if (element.type !== 'shape' || !hasShapeProperties(element)) return undefined;
	return element.pathData ? 'custom' : getShapeType(element.shapeType);
}

/**
 * Whether two shapes are different geometry families, which PowerPoint never
 * pairs on identity alone (only a `!!` name pairs them).
 */
export function morphGeometryBreaksPair(from: PptxElement, to: PptxElement): boolean {
	const a = morphGeometryKey(from);
	const b = morphGeometryKey(to);
	return a !== undefined && b !== undefined && a !== b;
}

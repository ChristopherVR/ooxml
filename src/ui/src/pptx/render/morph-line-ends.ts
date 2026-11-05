/**
 * What a line looks like at its ENDS: arrowheads, and how a morph treats a
 * change in them.
 *
 * Measured on PowerPoint 16 (`CreateVideo`, 60 fps, 720p, Morph by object): a
 * connector whose arrowhead changes type or size (triangle to oval, medium to
 * long) keeps its body at full strength and dissolves one head into the other
 * in place (no scaling of the head), and a dashed line turning into another
 * dash pattern blends the two. But a line that GAINS or LOSES something at an
 * end or along its length (an arrowhead at either end appearing or going, a
 * dashed line turning solid or back) is NOT paired: the old line fades out
 * where it stood (body included, gone by about 15-25% of the morph) and the new
 * one fades in from about 50-55%, the unmatched pair's dissolve (compare
 * `morph-geometry-key`), measured over eight such changes in two decks. So the
 * arrowheads and the dash are part of the pair's appearance, like its fill.
 *
 * @module render/morph-line-ends
 */
import type { ConnectorArrowType, PptxElement } from 'ooxml-core/pptx';

import { normalizeArrow } from './connector-markers';

/** The normalised arrowheads of a connector style: no head is `''`, a default size is `med`. */
export function morphArrowSignature(style: object | undefined): string {
	const record = (style ?? {}) as Record<string, unknown>;
	return ['Start', 'End']
		.map((end) => {
			const type = normalizeArrow(record[`connector${end}Arrow`] as ConnectorArrowType | undefined);
			if (!type) return '';
			const width = record[`connector${end}ArrowWidth`] ?? 'med';
			const length = record[`connector${end}ArrowLength`] ?? 'med';
			return `${type}/${String(width)}/${String(length)}`;
		})
		.join('>');
}

/** Whether a connector paints any dash at all (`solid` or none is a plain stroke). */
const dashed = (style: Record<string, unknown>): boolean =>
	style.strokeDash !== undefined && style.strokeDash !== 'solid';

/**
 * Whether two connectors differ in having a dash or an arrowhead at an end,
 * which PowerPoint does not pair on identity alone (a `!!` name still does).
 */
export function morphLineEndsBreakPair(from: PptxElement, to: PptxElement): boolean {
	if (from.type !== 'connector' || to.type !== 'connector') return false;
	const a = ((from as { shapeStyle?: object }).shapeStyle ?? {}) as Record<string, unknown>;
	const b = ((to as { shapeStyle?: object }).shapeStyle ?? {}) as Record<string, unknown>;
	if (dashed(a) !== dashed(b)) return true;
	return ['Start', 'End'].some(
		(end) =>
			Boolean(normalizeArrow(a[`connector${end}Arrow`] as ConnectorArrowType | undefined)) !==
			Boolean(normalizeArrow(b[`connector${end}Arrow`] as ConnectorArrowType | undefined)),
	);
}

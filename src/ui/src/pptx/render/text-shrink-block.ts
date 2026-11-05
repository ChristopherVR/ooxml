/**
 * The horizontal offset of an `a:bodyPr/@anchorCtr="1"` text block.
 *
 * PowerPoint shrinks the block to its longest line (the bounding box of every
 * wrapped line) and centres THAT box in the text area; each line keeps its own
 * alignment inside it (PowerPoint 16 exports: left, centre and right aligned,
 * one to four wrapped lines, insets and a turned box). CSS `fit-content` only
 * shrinks a block that does not wrap: a wrapped one keeps the full column, so
 * its lines sat at the wrong x. The lines are broken here with the canvas
 * measurer the morph text layout already uses, and the whole block is moved by
 * the one offset that centres their bounding box.
 *
 * @module render/text-shrink-block
 */
import type { PptxElement } from 'ooxml-core/pptx';

import { measureMorphTextWithCanvas, readFrame } from './morph-text-frame-read';
import { layoutMorphTextWords } from './morph-text-lines';

/**
 * How far (px, positive to the right) the full-column block has to move to be
 * the shrink-wrapped block PowerPoint draws, or `undefined` when the lines
 * cannot be measured with confidence (no canvas, mixed faces, bullets, ...),
 * where the caller keeps `fit-content`.
 */
export function shrinkBlockOffset(element: PptxElement): number | undefined {
	const frame = readFrame(element, { rotated: true, withBody: true, flippedH: true });
	if (!frame?.anchorCenter) {
		return undefined;
	}
	const layout = layoutMorphTextWords(
		{ ...frame, anchorCenter: false },
		measureMorphTextWithCanvas,
	);
	if (!layout) {
		return undefined;
	}
	const ruled = layout.lines.filter((line) => line.x1 > line.x0);
	if (!ruled.length) {
		return undefined;
	}
	const left = Math.min(...ruled.map((line) => line.x0));
	const right = Math.max(...ruled.map((line) => line.x1));
	const column = frame.insetLeft + frame.available / 2;
	return Math.round((column - (left + right) / 2) * 100) / 100;
}

import type { CanvasSize, InsertChartKind, InsertChartOptions } from 'ooxml-ui/pptx';
import { createDefaultChartElement } from 'ooxml-ui/pptx';
import type { PptxChartType, PptxElement } from 'pptx-viewer-core';

import { centerOnCanvas } from './editor-insert';

/**
 * Pure factory for the Insert > Chart action: wraps the shared
 * `createDefaultChartElement` (fully-populated sample data, ready to render)
 * and centres the result on the slide canvas.
 */

/**
 * Build a new, centred chart element for the given insert-dropdown entry.
 * `'column'` yields vertical columns, `'bar'` horizontal bars; a raw
 * `PptxChartType` is still accepted for API callers. `options` carries the
 * Insert Chart dialog's title and grouping.
 */
export function buildChartInsertElement(
	chartKind: InsertChartKind | PptxChartType,
	canvasSize: CanvasSize,
	options?: InsertChartOptions,
): PptxElement {
	const el: PptxElement = createDefaultChartElement(chartKind, undefined, options);
	centerOnCanvas(el, canvasSize);
	return el;
}

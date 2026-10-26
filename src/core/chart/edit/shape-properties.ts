import type { DrawingFill } from '../../drawingml/types';
import type { ChartShapeProperties } from '../model-series';
import { editChartLine, type ChartLineEdit } from './line-options';

export interface ChartShapeEdit {
	fill?: DrawingFill | null;
	line?: ChartLineEdit | null;
}

/** Sparse shape edits retain geometry, effects, theme references and unknown source children. */
export function editChartShapeProperties(
	current: ChartShapeProperties | undefined,
	edit: ChartShapeEdit,
): ChartShapeProperties {
	const result = structuredClone(current ?? {});
	if (edit.fill === null) delete result.fill;
	else if (edit.fill !== undefined) result.fill = structuredClone(edit.fill);
	if (edit.line === null) delete result.line;
	else if (edit.line !== undefined) result.line = editChartLine(result.line, edit.line);
	return result;
}

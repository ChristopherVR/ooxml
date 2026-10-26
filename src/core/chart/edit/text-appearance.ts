import { NS, first, type XmlElement } from '../../xml';
import { parseDrawingFill, parseDrawingLine } from '../../drawingml/drawing-fill';
import { fillAppearancePatch, lineAppearancePatch } from './appearance-fill';
import { editChartLine } from './line-options';
import type { ChartTextEdit } from './text';

/** Resolve host differences separately for each run, retaining its own authored theme tokens. */
export function resolveChartTextPaintEdit(node: XmlElement, edit: ChartTextEdit): ChartTextEdit {
	let next = edit;
	if (edit.fillAppearance)
		next = {
			...next,
			fill:
				fillAppearancePatch(
					parseDrawingFill(node),
					edit.fillAppearance.before,
					edit.fillAppearance.after,
				) ?? null,
		};
	if (edit.lineAppearance) {
		const current = parseDrawingLine(first(node, 'ln', NS.a));
		next = {
			...next,
			line: editChartLine(
				current,
				lineAppearancePatch(current, edit.lineAppearance.before, edit.lineAppearance.after),
			),
		};
	}
	return next;
}

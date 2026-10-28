import type { DrawingColor } from '../drawingml/types';
import type { DiagramColorStyleLabel, DiagramDrawingShape } from './types';

/** Apply authored colour-list roles, rather than the engine's category preview palette. */
export function applyDiagramRoleColors(
	shape: DiagramDrawingShape,
	role: DiagramColorStyleLabel | undefined,
	index: number,
	minorFont?: string,
): void {
	const at = (colors: readonly DrawingColor[]) => colors[index % colors.length];
	const fill = role && at(role.fill.colors);
	if (fill && shape.fill?.kind !== 'none' && shape.fill?.kind !== 'picture')
		shape.fill = { kind: 'solid', color: structuredClone(fill) };
	const line = role && at(role.line.colors);
	if (line && shape.line) shape.line.fill = { kind: 'solid', color: structuredClone(line) };
	const text = role && at(role.textFill.colors);
	for (const paragraph of shape.text?.paragraphs ?? []) {
		for (const run of paragraph.runs) {
			if (!run.color && text) run.color = structuredClone(text);
			if (!run.typeface && minorFont) run.typeface = minorFont;
		}
	}
}

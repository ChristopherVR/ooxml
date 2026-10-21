import type { VisioEdit } from 'ooxml-core/visio';
import { visioWithCalloutLeaders, visioWithDataGraphics } from 'ooxml-core/visio/ui';
import type { ViewerState } from './controller';

/**
 * The edits that delete the current selection, with the generated graphics and callout leaders
 * that belong to it, and the sentence to announce. Undefined when nothing can be deleted.
 */
export function planDeleteSelection(
	state: ViewerState,
): { edits: VisioEdit[]; message: string } | undefined {
	const shapes = state.selectedShapes;
	const page = state.document?.pages[state.pageIndex];
	if (!shapes.length || !page || shapes.some((shape) => shape.pageId && shape.pageId !== page.id))
		return undefined;
	return {
		edits: visioWithCalloutLeaders(
			page,
			visioWithDataGraphics(
				page,
				shapes.map((shape) => shape.id),
			),
		).map((shapeId) => ({ type: 'delete-shape', pageId: page.id, shapeId })),
		message:
			shapes.length === 1
				? `Deleted ${shapes[0]!.name || `shape ${shapes[0]!.id}`}.`
				: `Deleted ${shapes.length} shapes.`,
	};
}

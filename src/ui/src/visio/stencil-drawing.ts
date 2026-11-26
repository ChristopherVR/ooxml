import type { VisioStencilEdit } from 'ooxml-core/visio';
import type { ViewerController } from './controller';
import { documentMaster } from './shapes-document';

const SELECTION_REASON = 'Select one editable 2D shape in a drawing.';

export function stencilDrawing(controller: ViewerController) {
	const sourceAvailable = () => {
		const { edit, loading } = controller.state;
		return edit.sourceAvailable && !edit.busy && !loading;
	};
	const selectedShape = () => {
		const state = controller.state;
		const item = state.selectedShapes.length === 1 ? state.selectedShape : undefined;
		const page = state.document?.pages[state.pageIndex];
		const shape =
			page && item && (item.pageId === undefined || item.pageId === page.id)
				? page.shapes.find(({ id }) => id === item.id)
				: undefined;
		return page &&
			shape &&
			(shape.masterId || (shape.kind === 'shape' && !shape.layerIds?.length && !shape.lineEnds))
			? { pageId: page.id, shape }
			: undefined;
	};
	const captureMaster = (id: string): VisioStencilEdit => ({
		type: 'import-master',
		stencil: controller.exportVsdx().bytes,
		id,
	});
	return {
		sourceAvailable,
		documentMaster: (id: string) => documentMaster(controller.state.document, id),
		captureMaster,
		selectionReason: () =>
			controller.canEdit && sourceAvailable() && selectedShape() ? undefined : SELECTION_REASON,
		selection(name: string): VisioStencilEdit {
			const selected = selectedShape();
			if (!controller.canEdit || !sourceAvailable() || !selected) throw new Error(SELECTION_REASON);
			if (selected.shape.masterId) return captureMaster(selected.shape.masterId);
			return {
				type: 'add-shape',
				drawing: controller.exportVsdx().bytes,
				pageId: selected.pageId,
				shapeId: selected.shape.id,
				name,
			};
		},
	};
}

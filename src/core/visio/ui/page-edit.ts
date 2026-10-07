import type { VisioEdit } from '../edit-commands.js';
import type { VisioPage } from '../model.js';

/** Convert viewer page-inch controls to the edit API's drawing-inch coordinates.
 * Host calls to editVsdx/applyEdits keep their existing drawing-inch contract.
 */
export function visioPageEditToDrawing(page: VisioPage, edit: VisioEdit): VisioEdit {
	const ratio = page.drawingToPageScale ?? 1;
	if (!Number.isFinite(ratio) || ratio <= 0 || page.id !== edit.pageId)
		throw new Error('Invalid page scale or edit page.');
	if (ratio === 1) return edit;
	const coordinate = (value: number) => {
		const drawing = value / ratio;
		if (!Number.isFinite(drawing)) throw new Error('Drawing coordinate exceeds finite limits.');
		return drawing;
	};
	switch (edit.type) {
		case 'create-rectangle':
			return {
				...edit,
				x: coordinate(edit.x),
				y: coordinate(edit.y),
				width: coordinate(edit.width),
				height: coordinate(edit.height),
			};
		case 'move-shape':
			return { ...edit, x: coordinate(edit.x), y: coordinate(edit.y) };
		case 'resize-shape':
			return { ...edit, width: coordinate(edit.width), height: coordinate(edit.height) };
		default:
			return edit;
	}
}

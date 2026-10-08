import { fail } from './package-common';
import { snapshotDuplicateShapes } from './edit-duplicate-commands';
import { snapshotVisioClipboard, type VisioClipboardSnapshot } from './clipboard-types';

/** Drawing-inch offsets are explicit; cross-scale clipboard conversion is not yet supported. */
export interface VisioPasteShapesEdit {
	type: 'paste-shapes';
	pageId: string;
	clipboard: VisioClipboardSnapshot;
	copies: readonly { shapeId: string; newShapeId: string }[];
	offsetX: number;
	offsetY: number;
}
export function snapshotPasteShapes(edit: VisioPasteShapesEdit): VisioPasteShapesEdit {
	const clipboard = snapshotVisioClipboard(edit.clipboard);
	const copies = snapshotDuplicateShapes({
		type: 'duplicate-shapes',
		pageId: edit.pageId,
		copies: edit.copies,
		offsetX: edit.offsetX,
		offsetY: edit.offsetY,
	});
	if (
		copies.copies.length !== clipboard.shapes.length ||
		copies.copies.some((copy) => !clipboard.selectionIds.includes(copy.shapeId))
	)
		fail('INVALID_CLIPBOARD', 'Paste mappings must cover every captured shape exactly once.');
	return {
		type: 'paste-shapes',
		pageId: copies.pageId,
		copies: copies.copies,
		offsetX: copies.offsetX,
		offsetY: copies.offsetY,
		clipboard,
	};
}

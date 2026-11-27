import { MAX_SHARED_CELL_PICTURE_BYTES, cellView, getCell } from 'ooxml-core/xlsx';
import {
	editing,
	icon,
	sheetLocked,
	target,
	type Command,
	type EditorContext,
} from 'ooxml-core/xlsx/ui';

type Target = NonNullable<ReturnType<typeof target>>;

/**
 * In-cell pictures are shared in collaboration. Floating pictures are not, so turning an in-cell
 * picture into a floating one waits until sharing ends (collaborators would lose the picture).
 */
const floatingAvailable = (ctx: EditorContext): boolean => !ctx.isCollaborating?.();

function floatingUnavailable(other: string): (ctx: EditorContext) => string {
	return (ctx) =>
		floatingAvailable(ctx)
			? other
			: 'Floating pictures are not shared yet. Stop sharing to convert this picture.';
}

const INVALID_PICTURE = 'Use a PNG, JPEG or GIF picture within the size limits.';
const TOO_LARGE_TO_SHARE = 'Pictures in cells are shared up to 4 MB. Choose a smaller picture.';

/** While shared, a picture over the room's limit would reach collaborators as #VALUE! only. */
function refusedWhileShared(ctx: EditorContext, size: number): boolean {
	if (!ctx.isCollaborating?.() || size <= MAX_SHARED_CELL_PICTURE_BYTES) return false;
	ctx.toast(ctx.t(TOO_LARGE_TO_SHARE), 'warning');
	return true;
}

export async function insertCellPicture(
	ctx: EditorContext,
	current: Target,
	file: File,
): Promise<void> {
	if (file.size > 8 * 1024 * 1024) {
		ctx.toast(ctx.t(INVALID_PICTURE), 'warning');
		return;
	}
	if (refusedWhileShared(ctx, file.size)) return;
	const bytes = new Uint8Array(await file.arrayBuffer());
	if (
		ctx.readOnly() ||
		ctx.session() !== current.session ||
		target(ctx)?.ws !== current.ws ||
		sheetLocked(ctx)
	)
		return;
	if (!current.session.addCellPicture(current.sheet, current.active, bytes))
		ctx.toast(ctx.t(INVALID_PICTURE), 'warning');
}

function selectedCellPicture(ctx: EditorContext) {
	const current = target(ctx);
	if (!current || ctx.selection.get().drawing !== undefined) return undefined;
	const { row, col } = current.active;
	const cell = getCell(current.ws, row, col);
	if (
		!cell ||
		cell.formula !== undefined ||
		(!cell.inCellPicture && cell.valueMetadata === undefined)
	)
		return undefined;
	const picture = cellView(current.session.workbook, current.sheet, row, col).picture;
	return picture ? { current, picture } : undefined;
}

function selectedDrawingPicture(ctx: EditorContext) {
	const current = target(ctx);
	const index = ctx.selection.get().drawing;
	return current && index !== undefined && current.ws.drawings[index]?.kind === 'image'
		? { current, index }
		: undefined;
}

export function cellPictureCommands(): Command[] {
	return [
		editing({
			id: 'picture.place-in-cell',
			label: 'Convert to In-Cell Picture',
			unavailable: 'Select a floating picture on an editable sheet.',
			icon: icon('picture'),
			enabled: (ctx) => Boolean(selectedDrawingPicture(ctx)),
			run: (ctx) => {
				const selected = selectedDrawingPicture(ctx);
				if (!selected) return;
				const { current, index } = selected;
				const drawing = current.ws.drawings[index];
				const bytes =
					drawing?.kind === 'image'
						? current.session.workbook.source?.parts.get(drawing.partName)
						: undefined;
				if (refusedWhileShared(ctx, bytes?.length ?? 0)) return;
				const at = current.session.placePictureInCell(current.sheet, index);
				if (at)
					ctx.selection.set({
						active: at,
						anchor: at,
						ranges: [{ start: at, end: at }],
						drawing: undefined,
					});
				else ctx.toast(ctx.t(INVALID_PICTURE), 'warning');
			},
		}),
		editing({
			id: 'picture.place-over-cells',
			label: 'Convert to Floating Picture',
			unavailable: floatingUnavailable('Select an in-cell picture on an editable sheet.'),
			icon: icon('picture'),
			enabled: (ctx) => floatingAvailable(ctx) && Boolean(selectedCellPicture(ctx)),
			run: (ctx) => {
				const selected = selectedCellPicture(ctx);
				if (!selected) return;
				const { current } = selected;
				const index = current.session.placePictureOverCells(current.sheet, current.active);
				if (index !== undefined) ctx.selection.set({ drawing: index });
			},
		}),
		editing({
			id: 'picture.cell-alt-text',
			label: 'Alt Text',
			unavailable: 'Select an in-cell picture on an editable sheet.',
			enabled: (ctx) => Boolean(selectedCellPicture(ctx)),
			run: async (ctx) => {
				const selected = selectedCellPicture(ctx);
				if (!selected) return;
				const { editPictureAltText } = await import('../dialogs/rich-values');
				await editPictureAltText(ctx, selected.current, selected.picture.altText);
			},
		}),
	];
}

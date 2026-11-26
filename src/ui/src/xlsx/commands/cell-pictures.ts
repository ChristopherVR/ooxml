import { cellView, getCell } from 'ooxml-core/xlsx';
import {
	editing,
	icon,
	sheetLocked,
	target,
	type Command,
	type EditorContext,
} from 'ooxml-core/xlsx/ui';

type Target = NonNullable<ReturnType<typeof target>>;
const INVALID_PICTURE = 'Use a PNG, JPEG or GIF picture within the size limits.';

export async function insertCellPicture(
	ctx: EditorContext,
	current: Target,
	file: File,
): Promise<void> {
	if (file.size > 8 * 1024 * 1024) {
		ctx.toast(ctx.t(INVALID_PICTURE), 'warning');
		return;
	}
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
			unavailable: 'Select an in-cell picture on an editable sheet.',
			icon: icon('picture'),
			enabled: (ctx) => Boolean(selectedCellPicture(ctx)),
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

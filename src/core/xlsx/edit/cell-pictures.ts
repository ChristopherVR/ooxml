import { MAX_COL, MAX_ROW, type CellAddress } from '../address';
import { getCell } from '../cells';
import { pictureAnchorAt } from '../layout/anchors';
import { cellError, type ImageObject } from '../model';
import { cellPictureInfo, embeddedPicture, type EmbeddedCellPicture } from '../rich-values';
import { clearContents } from './cell-values';
import { newMediaPart } from './charts';
import { ensureCell, pruneCell, sheetAt, sheetDrawingAt, type EditContext } from './context';

function checkAddress(at: CellAddress): void {
	if (
		!Number.isInteger(at.row) ||
		!Number.isInteger(at.col) ||
		at.row < 0 ||
		at.col < 0 ||
		at.row > MAX_ROW ||
		at.col > MAX_COL
	)
		throw new RangeError('Invalid picture cell.');
}

const rangeAt = (at: CellAddress) => ({ start: at, end: at });
const cellScope = (sheet: number, at: CellAddress) => ({
	kind: 'cells' as const,
	sheet,
	ranges: [rangeAt(at)],
});
const drawingScope = (sheet: number) => ({
	kind: 'parts' as const,
	sheet,
	parts: ['drawings' as const],
});

function place(
	ctx: EditContext,
	sheet: number,
	at: CellAddress,
	picture: EmbeddedCellPicture,
): void {
	const cell = ensureCell(sheetAt(ctx.workbook, sheet), at.row, at.col);
	clearContents(cell);
	cell.value = cellError('#VALUE!');
	cell.inCellPicture = picture;
}

export function addCellPicture(
	ctx: EditContext,
	s: number,
	at: CellAddress,
	bytes: Uint8Array,
	altText = '',
): boolean {
	checkAddress(at);
	sheetAt(ctx.workbook, s);
	const info = cellPictureInfo(bytes);
	if (!info) return false;
	const partName = newMediaPart(ctx.workbook, info.mimeType);
	ctx.run(
		'Insert picture in cell',
		'cells',
		[cellScope(s, at)],
		() => {
			ctx.workbook.source ??= { parts: new Map() };
			ctx.workbook.source.parts.set(partName, bytes.slice());
			place(ctx, s, at, { partName, altText: altText.slice(0, 1024) });
		},
		{ sheet: s, ranges: [rangeAt(at)] },
	);
	return true;
}

export function placePictureInCell(
	ctx: EditContext,
	s: number,
	index: number,
): CellAddress | undefined {
	const drawing = sheetDrawingAt(ctx.workbook, s, index);
	if (drawing?.kind !== 'image') return undefined;
	const { partName, anchor, description } = drawing;
	const bytes = ctx.workbook.source?.parts.get(partName);
	if (!bytes || !cellPictureInfo(bytes)) return undefined;
	const at = { row: anchor.from.row, col: anchor.from.col };
	checkAddress(at);
	const sheet = sheetAt(ctx.workbook, s);
	ctx.run(
		'Place picture in cell',
		'cells',
		[cellScope(s, at), drawingScope(s)],
		() => {
			place(ctx, s, at, { partName, altText: (description ?? '').slice(0, 1024) });
			sheet.drawings.splice(index, 1);
		},
		{ sheet: s, ranges: [rangeAt(at)] },
	);
	return at;
}

export function placePictureOverCells(
	ctx: EditContext,
	s: number,
	at: CellAddress,
): number | undefined {
	const sheet = sheetAt(ctx.workbook, s);
	const cell = getCell(sheet, at.row, at.col);
	const picture =
		cell && cell.formula === undefined ? embeddedPicture(ctx.workbook, cell) : undefined;
	const bytes = picture && ctx.workbook.source?.parts.get(picture.partName);
	const info = bytes && cellPictureInfo(bytes);
	if (!cell || !picture || !info) return undefined;
	const drawing: ImageObject = {
		kind: 'image',
		partName: picture.partName,
		contentType: info.mimeType,
		anchor: pictureAnchorAt(at, info.pixelWidth, info.pixelHeight),
		description: picture.altText,
	};
	return ctx.run(
		'Place picture over cells',
		'cells',
		[cellScope(s, at), drawingScope(s)],
		() => {
			clearContents(cell);
			pruneCell(sheet, at.row, at.col);
			return sheet.drawings.push(drawing) - 1;
		},
		{ sheet: s, ranges: [rangeAt(at)] },
	);
}

export function setCellPictureAltText(
	ctx: EditContext,
	s: number,
	at: CellAddress,
	text: string,
): boolean {
	const cell = getCell(sheetAt(ctx.workbook, s), at.row, at.col);
	const picture =
		cell && cell.formula === undefined ? embeddedPicture(ctx.workbook, cell) : undefined;
	if (!cell || !picture) return false;
	const altText = text.slice(0, 1024);
	if (altText === picture.altText) return true;
	ctx.run(
		'Edit picture alt text',
		'cells',
		[cellScope(s, at)],
		() => {
			delete cell.valueMetadata;
			cell.inCellPicture = { ...picture, altText };
		},
		{ sheet: s, ranges: [rangeAt(at)] },
	);
	return true;
}

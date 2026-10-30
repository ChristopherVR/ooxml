import { attribute, children } from './sheet';
import { fail } from './package-common';
import { cells, numeric, setCell } from './edit-geometry-cells';
import { recalculateVisioCells, type VisioCellKey } from './edit-recalculate';

/** Visio's stored neutral value for Brightness and Contrast (0%). */
export const VISIO_PICTURE_NEUTRAL = 0.5;

/**
 * Picture Format > Adjust: Brightness and Contrast are the shape cells `Brightness` and `Contrast`
 * (0 to 1, 0.5 neutral), each optional so a preset changes one of them.
 */
export interface VisioPictureAdjustEdit {
	type: 'set-picture-adjust';
	pageId: string;
	shapeId: string;
	brightness?: number;
	contrast?: number;
}
/**
 * Picture Format > Crop: `square` crops the longer side equally so the picture is square;
 * `reset` restores the whole picture (the frame becomes the picture's own size).
 */
export interface VisioPictureCropEdit {
	type: 'crop-picture';
	pageId: string;
	shapeId: string;
	crop: 'square' | 'reset';
}
export type VisioPictureEdit = VisioPictureAdjustEdit | VisioPictureCropEdit;

export const isVisioPictureEdit = (edit: { type: string }): edit is VisioPictureEdit =>
	edit.type === 'set-picture-adjust' || edit.type === 'crop-picture';

const unit = (value: unknown): number => {
	if (typeof value !== 'number' || !Number.isFinite(value) || value < 0 || value > 1)
		fail('INVALID_EDIT', 'Brightness and contrast are fractions from 0 to 1.');
	return value;
};

/** Copy and validate a picture edit; the picture itself is checked when the edit is applied. */
export function snapshotPictureEdit(edit: VisioPictureEdit): VisioPictureEdit {
	if (typeof edit.pageId !== 'string' || !edit.pageId || edit.pageId.length > 256)
		fail('INVALID_EDIT', 'Invalid edit target.');
	if (typeof edit.shapeId !== 'string' || !/^[1-9]\d{0,9}$/.test(edit.shapeId))
		fail('INVALID_EDIT', 'Picture shape IDs require canonical positive unsigned integers.');
	if (edit.type === 'crop-picture') {
		if (edit.crop !== 'square' && edit.crop !== 'reset')
			fail('INVALID_EDIT', 'Unknown crop command.');
		return { type: edit.type, pageId: edit.pageId, shapeId: edit.shapeId, crop: edit.crop };
	}
	if (edit.brightness === undefined && edit.contrast === undefined)
		fail('INVALID_EDIT', 'A picture adjustment sets brightness or contrast.');
	return {
		type: edit.type,
		pageId: edit.pageId,
		shapeId: edit.shapeId,
		...(edit.brightness === undefined ? {} : { brightness: unit(edit.brightness) }),
		...(edit.contrast === undefined ? {} : { contrast: unit(edit.contrast) }),
	};
}

function findShape(root: Element, shapeId: string): Element | undefined {
	return Array.from(root.getElementsByTagNameNS(root.namespaceURI, 'Shape')).find(
		(shape) => attribute(shape, 'ID') === shapeId,
	);
}

/** A picture is a Foreign shape whose ForeignData is an embedded bitmap. */
function pictureShape(root: Element, shapeId: string): Element {
	const shape = findShape(root, shapeId);
	const data = shape && children(shape, 'ForeignData')[0];
	if (!shape || attribute(data, 'ForeignType') !== 'Bitmap')
		fail('EDIT_TARGET_NOT_FOUND', 'A picture shape is required.');
	return shape;
}

/** Apply a picture edit to the page root. Returns the pages whose cells changed. */
export function applyPictureEdit(
	roots: ReadonlyMap<string, Element>,
	edit: VisioPictureEdit,
	check: () => void,
): readonly string[] {
	check();
	const root = roots.get(edit.pageId);
	if (!root) fail('EDIT_TARGET_NOT_FOUND', 'Page does not exist.');
	const shape = pictureShape(root, edit.shapeId);
	if (edit.type === 'set-picture-adjust') {
		if (edit.brightness !== undefined) setCell(shape, 'Brightness', edit.brightness);
		if (edit.contrast !== undefined) setCell(shape, 'Contrast', edit.contrast);
		return [edit.pageId];
	}
	if (numeric(cells(shape).get('Angle'), 0) !== 0)
		fail('UNSUPPORTED_GEOMETRY_EDIT', 'Crop needs an unrotated picture.');
	const width = numeric(cells(shape).get('Width'));
	const height = numeric(cells(shape).get('Height'));
	const imageWidth = numeric(cells(shape).get('ImgWidth'), width);
	const imageHeight = numeric(cells(shape).get('ImgHeight'), height);
	const offsetX = numeric(cells(shape).get('ImgOffsetX'), 0);
	const offsetY = numeric(cells(shape).get('ImgOffsetY'), 0);
	const left = numeric(cells(shape).get('PinX')) - numeric(cells(shape).get('LocPinX'), width / 2);
	const bottom =
		numeric(cells(shape).get('PinY')) - numeric(cells(shape).get('LocPinY'), height / 2);
	// The picture keeps its scale: the frame moves and shrinks while the picture stays put.
	const frame =
		edit.crop === 'reset'
			? {
					width: imageWidth,
					height: imageHeight,
					left: left + offsetX,
					bottom: bottom + offsetY,
					offsetX: 0,
					offsetY: 0,
				}
			: {
					width: width - 2 * Math.max(0, (width - height) / 2),
					height: height - 2 * Math.max(0, (height - width) / 2),
					left: left + Math.max(0, (width - height) / 2),
					bottom: bottom + Math.max(0, (height - width) / 2),
					offsetX: offsetX - Math.max(0, (width - height) / 2),
					offsetY: offsetY - Math.max(0, (height - width) / 2),
				};
	if (edit.crop === 'square' && width === height) return [];
	setCell(shape, 'Width', frame.width);
	setCell(shape, 'Height', frame.height);
	setCell(shape, 'LocPinX', frame.width / 2, 'Width*0.5');
	setCell(shape, 'LocPinY', frame.height / 2, 'Height*0.5');
	setCell(shape, 'PinX', frame.left + frame.width / 2);
	setCell(shape, 'PinY', frame.bottom + frame.height / 2);
	setCell(shape, 'ImgOffsetX', frame.offsetX);
	setCell(shape, 'ImgOffsetY', frame.offsetY);
	setCell(shape, 'ImgWidth', imageWidth);
	setCell(shape, 'ImgHeight', imageHeight);
	const changed: VisioCellKey[] = [
		'Width',
		'Height',
		'PinX',
		'PinY',
		'LocPinX',
		'LocPinY',
		'ImgOffsetX',
		'ImgOffsetY',
		'ImgWidth',
		'ImgHeight',
	].map((cell) => ({ pageId: edit.pageId, shapeId: edit.shapeId, cell }));
	return [...new Set([edit.pageId, ...recalculateVisioCells(roots, changed, { check })])];
}

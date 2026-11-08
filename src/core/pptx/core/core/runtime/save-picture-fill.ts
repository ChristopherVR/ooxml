/**
 * Picture fills and picture layout of NEW elements: the `a:blipFill` of a
 * shape filled with a picture (`ShapeStyle.fillMode === 'image'`), and the
 * `a:tile` / `a:stretch/a:fillRect` of a new picture.
 *
 * The `a:blipFill` itself is built by the shared DrawingML picture-fill
 * writer (`drawingFillXml`, `applyPictureLayout`), which owns the attribute
 * forms and the range checks, and is then turned into this area's XML object
 * form. Loaded pictures and loaded picture fills never come through here:
 * they keep their authored `a:blipFill`.
 *
 * @module save-picture-fill
 */

import type { DrawingPictureLayout } from '../../../../drawingml/picture-layout';
import { drawingFillXml } from '../../../../drawingml/write-fill';
import type { PptxImageLayout, ShapeStyle, XmlObject } from '../../types';
import { setFillChoice } from './fill-choice-group';

/** What the writers need from the save runtime. */
export interface PictureFillSaveContext {
	/** EMU per CSS pixel (9525). */
	readonly emuPerPx: number;
	/** Parse an XML string into this area's XML object form. */
	parse(xml: string): XmlObject;
	/** Add `dataUrl` as a media part of the slide; its relationship id, or `undefined`. */
	embed(dataUrl: string): string | undefined;
	/** Report a compatibility warning for the element being saved. */
	warn(code: string, message: string): void;
}

const has = (value: unknown): boolean => value !== undefined;

/**
 * The shared writer's layout for the model's fields, or `undefined` when
 * they say nothing (the default stretch over the whole frame). Tile fields
 * win over the fill rectangle, as on load.
 */
export function drawingPictureLayout(
	layout: PptxImageLayout,
	emuPerPx: number,
): DrawingPictureLayout | undefined {
	const offset = (px: number | undefined) =>
		px === undefined ? undefined : Math.round(px * emuPerPx);
	if (
		[
			layout.tileOffsetX,
			layout.tileOffsetY,
			layout.tileScaleX,
			layout.tileScaleY,
			layout.tileFlip,
			layout.tileAlignment,
		].some(has)
	) {
		const tile: NonNullable<DrawingPictureLayout['tile']> = {};
		const offsetX = offset(layout.tileOffsetX);
		const offsetY = offset(layout.tileOffsetY);
		if (offsetX !== undefined) tile.offsetXEmu = offsetX;
		if (offsetY !== undefined) tile.offsetYEmu = offsetY;
		if (layout.tileScaleX !== undefined) tile.scaleX = layout.tileScaleX;
		if (layout.tileScaleY !== undefined) tile.scaleY = layout.tileScaleY;
		if (layout.tileFlip !== undefined) tile.flip = layout.tileFlip;
		if (layout.tileAlignment !== undefined) tile.alignment = layout.tileAlignment;
		return { tile };
	}
	if (
		[layout.fillRectLeft, layout.fillRectTop, layout.fillRectRight, layout.fillRectBottom].some(has)
	) {
		return {
			fillRect: {
				l: layout.fillRectLeft ?? 0,
				t: layout.fillRectTop ?? 0,
				r: layout.fillRectRight ?? 0,
				b: layout.fillRectBottom ?? 0,
			},
		};
	}
	return undefined;
}

/**
 * `a:blipFill` for relationship `relId`, written by the shared DrawingML
 * writer.
 *
 * @throws RangeError from the shared writer for a layout value the file
 *   cannot hold (a scale at or below 0, an unknown alignment or flip).
 */
function blipFillXml(
	relId: string,
	layout: DrawingPictureLayout,
	parse: PictureFillSaveContext['parse'],
): XmlObject {
	const xml = drawingFillXml({ kind: 'picture', relId, layout }) ?? '';
	const node = parse(xml)['a:blipFill'] as XmlObject;
	delete node['@_xmlns:a'];
	delete node['@_xmlns:r'];
	return node;
}

/** The model's layout, or the default stretch with a warning when the file cannot hold it. */
function checkedBlipFill(
	relId: string,
	layout: DrawingPictureLayout,
	ctx: PictureFillSaveContext,
): XmlObject {
	try {
		return blipFillXml(relId, layout, ctx.parse);
	} catch (error) {
		if (!(error instanceof RangeError)) {
			throw error;
		}
		ctx.warn(
			'SAVE_PICTURE_LAYOUT_INVALID',
			`Picture tile or fill rectangle could not be written (${error.message}); the picture is stretched over its frame instead.`,
		);
		const { rotWithShape } = layout;
		return blipFillXml(relId, rotWithShape === undefined ? {} : { rotWithShape }, ctx.parse);
	}
}

/**
 * Write the `a:blipFill` of a new shape whose fill is a picture, embedding
 * the picture as a media part of the slide. A picture that cannot be
 * embedded (not a data URL) leaves the shape unfilled with a warning.
 * `rotWithShape` defaults to `1`, as PowerPoint writes picture fills.
 */
export function writeNewShapePictureFill(
	spPr: XmlObject,
	style: ShapeStyle,
	ctx: PictureFillSaveContext,
): void {
	const relId = style.fillImageUrl ? ctx.embed(style.fillImageUrl) : undefined;
	if (!relId) {
		ctx.warn(
			'SAVE_SHAPE_PICTURE_FILL_UNSUPPORTED',
			'Shape picture fill could not be embedded (it needs a data: URL); the shape was saved without a fill.',
		);
		return;
	}
	const layout: DrawingPictureLayout = {
		...drawingPictureLayout(style.fillImageLayout ?? {}, ctx.emuPerPx),
		rotWithShape: style.fillImageRotWithShape ?? true,
	};
	setFillChoice(spPr, 'a:blipFill', checkedBlipFill(relId, layout, ctx));
}

/**
 * Replace the default `a:stretch` of a new picture's `p:blipFill` with the
 * tile or fill rectangle its model asks for. Nothing changes without one.
 */
export function applyNewPictureLayout(
	blipFill: XmlObject,
	layout: PptxImageLayout,
	ctx: PictureFillSaveContext,
): void {
	const drawingLayout = drawingPictureLayout(layout, ctx.emuPerPx);
	const relId = (blipFill['a:blip'] as XmlObject | undefined)?.['@_r:embed'];
	if (!drawingLayout || typeof relId !== 'string') {
		return;
	}
	const built = checkedBlipFill(relId, drawingLayout, ctx);
	delete blipFill['a:stretch'];
	delete blipFill['a:tile'];
	for (const key of ['a:tile', 'a:stretch']) {
		if (built[key] !== undefined) {
			blipFill[key] = built[key];
		}
	}
}

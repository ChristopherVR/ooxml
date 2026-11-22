import { patchPictureEffects, type PictureEffectPatch } from '../../drawingml/picture-effects';
import type { DrawingAnchor, DrawingObject } from '../model';
import { removeSlicers } from '../slicer/book';
import { deleteTimeline } from './timelines';
import { type EditContext, sheetAt, sheetDrawingAt } from './context';
import {
	normalizeWordArtGradient,
	wordArtPaintEquals,
	wordArtRgb,
	type WordArtText,
} from '../../drawingml/word-art';
import { wordArtShape } from '../word-art';

function drawingAt(ctx: EditContext, s: number, index: number): DrawingObject {
	const drawing = sheetDrawingAt(ctx.workbook, s, index);
	if (!drawing) throw new RangeError(`No drawing at index ${index} on sheet ${s}`);
	return drawing;
}

export function checkAnchor(anchor: DrawingAnchor): void {
	const points = [anchor.from, ...(anchor.to ? [anchor.to] : [])];
	for (const p of points)
		if (
			![p.row, p.col, p.rowOffset, p.colOffset].every(Number.isFinite) ||
			p.row < 0 ||
			p.col < 0 ||
			p.rowOffset < 0 ||
			p.colOffset < 0
		)
			throw new RangeError('A drawing anchor needs non-negative cell positions and offsets.');
	if (anchor.ext && !(anchor.ext.cx >= 0 && anchor.ext.cy >= 0))
		throw new RangeError('A drawing extent cannot be negative.');
}

function normalizeWordArt(value: WordArtText): WordArtText {
	const text = value.text.trim();
	const fill = wordArtRgb(value.style.fill);
	const outline = wordArtRgb(value.style.outline);
	const shadow = wordArtRgb(value.style.shadow);
	const gradient = normalizeWordArtGradient(value.style.gradient);
	if (
		!text ||
		!fill ||
		(value.style.gradient !== undefined && !gradient) ||
		(value.style.outline !== undefined && !outline) ||
		(value.style.shadow !== undefined && !shadow)
	) {
		throw new RangeError('WordArt requires text and valid colors.');
	}
	return {
		text,
		style: {
			fill,
			...(gradient ? { gradient } : {}),
			...(outline ? { outline } : {}),
			...(shadow ? { shadow } : {}),
			...(value.style.warp ? { warp: structuredClone(value.style.warp) } : {}),
		},
	};
}

/** Inserts native worksheet WordArt as one drawing edit. */
export function addWordArt(
	ctx: EditContext,
	s: number,
	anchor: DrawingAnchor,
	value: WordArtText,
): number {
	checkAnchor(anchor);
	const sheet = sheetAt(ctx.workbook, s);
	const shape = wordArtShape(normalizeWordArt(value), anchor);
	return ctx.run(
		'Insert WordArt',
		'annotations',
		[{ kind: 'sheet', sheet: s }],
		() => {
			sheet.drawings.push({
				kind: 'unsupported',
				description: 'text box',
				anchor: structuredClone(anchor),
				shape,
			});
			return sheet.drawings.length - 1;
		},
		{ sheet: s },
	);
}

export function editWordArt(ctx: EditContext, s: number, index: number, value: WordArtText): void {
	const drawing = drawingAt(ctx, s, index);
	if (drawing.kind !== 'unsupported' || !drawing.shape?.wordArt) return;
	const next = normalizeWordArt(value);
	const shape = drawing.shape;
	if (
		shape.text?.text === next.text &&
		shape.wordArt &&
		wordArtPaintEquals(shape.wordArt, next.style) &&
		JSON.stringify(shape.wordArt.warp) === JSON.stringify(next.style.warp)
	)
		return;
	ctx.run(
		'Edit WordArt',
		'annotations',
		[{ kind: 'sheet', sheet: s }],
		() => {
			const wordArt = wordArtShape(next, drawing.anchor);
			drawing.shape = { ...shape, wordArt: next.style, text: wordArt.text! };
			if (drawing.sourceXml) drawing.wordArtEdited = true;
		},
		{ sheet: s },
	);
}

/**
 * Moves or resizes a picture, chart or shape (one undo step). The saved drawing part is
 * regenerated from the model, so the new anchor is what Excel shows.
 */
export function setDrawingAnchor(
	ctx: EditContext,
	s: number,
	index: number,
	anchor: DrawingAnchor,
): void {
	const drawing = drawingAt(ctx, s, index);
	checkAnchor(anchor);
	ctx.run(
		'Move object',
		'annotations',
		[{ kind: 'sheet', sheet: s }],
		() => {
			drawing.anchor = structuredClone(anchor);
		},
		{ sheet: s },
	);
}

/** Deletes a picture, chart or shape (one undo step). */
export function deleteDrawing(ctx: EditContext, s: number, index: number): void {
	const doomed = drawingAt(ctx, s, index);
	const sheet = sheetAt(ctx.workbook, s);
	if (doomed.kind === 'timeline') {
		deleteTimeline(ctx, doomed.name);
		return;
	}
	if (doomed.kind === 'slicer') {
		ctx.run(
			'Delete slicer',
			'annotations',
			[{ kind: 'workbook' }],
			() => removeSlicers(ctx.workbook, (slicer) => slicer.name === doomed.name),
			{ sheet: s },
		);
		return;
	}
	ctx.run(
		'Delete object',
		'annotations',
		[{ kind: 'sheet', sheet: s }],
		() => {
			sheet.drawings.splice(index, 1);
		},
		{ sheet: s },
	);
}

/** Crops the picture in source-image fractions (one undo step). Negative insets pad the frame. */
export function setPictureCrop(
	ctx: EditContext,
	s: number,
	index: number,
	crop: import('../../drawingml/picture-crop').DrawingPictureCrop,
): void {
	const drawing = drawingAt(ctx, s, index);
	if (drawing.kind !== 'image') throw new RangeError('Only pictures can be cropped');
	const sides = ['cropLeft', 'cropTop', 'cropRight', 'cropBottom'] as const;
	if (
		sides.some((side) => crop[side] !== undefined && !Number.isFinite(crop[side])) ||
		1 - (crop.cropLeft ?? 0) - (crop.cropRight ?? 0) <= 0 ||
		1 - (crop.cropTop ?? 0) - (crop.cropBottom ?? 0) <= 0
	)
		throw new RangeError('Picture crop needs a positive source span and finite insets');
	ctx.run(
		'Crop picture',
		'annotations',
		[{ kind: 'sheet', sheet: s }],
		() => {
			drawing.crop = Object.fromEntries(
				sides.flatMap((side) => (crop[side] === undefined ? [] : [[side, crop[side]]])),
			);
		},
		{ sheet: s },
	);
}

/** Applies admitted picture effects in one undo step, retaining every unknown effect. */
export function editPictureEffects(
	ctx: EditContext,
	s: number,
	index: number,
	patch: PictureEffectPatch,
): void {
	const drawing = drawingAt(ctx, s, index);
	if (drawing.kind !== 'image') throw new RangeError('Only pictures have picture effects');
	const effects = patchPictureEffects(drawing.effects ?? [], patch);
	ctx.run(
		'Format picture',
		'annotations',
		[{ kind: 'sheet', sheet: s }],
		() => {
			drawing.effects = effects;
		},
		{ sheet: s },
	);
}

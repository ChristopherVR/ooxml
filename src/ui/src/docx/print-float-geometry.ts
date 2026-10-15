import type { PicturePlacement } from 'ooxml-core/docx';

/** A floating object's rectangle on its page, in CSS pixels. */
export interface FloatBox {
	x: number;
	y: number;
	width: number;
	height: number;
}

export type FloatHandle = 'nw' | 'n' | 'ne' | 'e' | 'se' | 's' | 'sw' | 'w';

export const FLOAT_HANDLES: readonly FloatHandle[] = ['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w'];
export const MIN_FLOAT_SIZE = 16;
/** One arrow press moves a selected float by this much (Ctrl+arrow moves 1px). */
export const NUDGE_PX = 6;

/**
 * Resizes `start` by dragging `handle` by (`dx`, `dy`). Corner handles keep the aspect ratio when
 * `lockAspect` is set (pictures), averaging the two axes the way Word's proportional resize
 * feels; edge handles change one dimension. The opposite edge or corner stays fixed.
 */
export function resizeBox(
	start: FloatBox,
	handle: FloatHandle,
	dx: number,
	dy: number,
	lockAspect: boolean,
): FloatBox {
	const sx = handle.includes('e') ? 1 : handle.includes('w') ? -1 : 0;
	const sy = handle.includes('s') ? 1 : handle.includes('n') ? -1 : 0;
	let width = start.width + sx * dx;
	let height = start.height + sy * dy;
	if (lockAspect && sx !== 0 && sy !== 0) {
		const floor = Math.max(MIN_FLOAT_SIZE / start.width, MIN_FLOAT_SIZE / start.height);
		const scale = Math.max(floor, (width / start.width + height / start.height) / 2);
		width = start.width * scale;
		height = start.height * scale;
	}
	width = Math.max(MIN_FLOAT_SIZE, width);
	height = Math.max(MIN_FLOAT_SIZE, height);
	return {
		x: sx < 0 ? start.x + start.width - width : start.x,
		y: sy < 0 ? start.y + start.height - height : start.y,
		width,
		height,
	};
}

export function moveBox(start: FloatBox, dx: number, dy: number): FloatBox {
	return { ...start, x: start.x + dx, y: start.y + dy };
}

/**
 * The placement that puts a float at `box`: offsets from the frame origin, with any alignment
 * dropped (the offset now says where it is). Wrapping, anchoring and frames are kept.
 */
export function placementAt(
	placement: PicturePlacement | undefined,
	box: FloatBox,
	origin: { x: number; y: number },
): PicturePlacement {
	const { align: _align, alignV: _alignV, ...rest } = placement ?? { wrap: 'none' as const };
	return {
		...rest,
		offsetXPx: Math.round(box.x - origin.x),
		offsetYPx: Math.round(box.y - origin.y),
	};
}

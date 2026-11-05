/**
 * Where a group child really sits in the stage, for the device-pixel stroke snap.
 *
 * A child's `x` / `y` are relative to its group's box, and a group's own `x` / `y`
 * are rounded to whole pixels while the children keep the exact (unrounded)
 * offset from the exact group origin. The snap rounds ABSOLUTE stage positions,
 * so it needs the chain of ancestor groups: the position the child is painted at
 * (the sum of the ancestors' model `x` / `y`) and the position PowerPoint puts it
 * at (the sum of their exact origins, from `xEmu` / `yEmu` where those still
 * describe the model).
 *
 * Bindings render a group's children from the group element, and every one of
 * them calls {@link getGroupChildParentFill} there, which calls
 * {@link registerGroupChildren}: the child to group links are kept in a WeakMap,
 * so no binding has to thread a parent through its renderer props. A child that
 * was never registered reads as top level (the previous behaviour).
 *
 * A rotated or flipped ancestor group turns the child's axes: its strokes are
 * anti-aliased in PowerPoint's export too and are not snapped ({@link snapPlacement}
 * returns `undefined`).
 *
 * @module render/device-pixel-group
 */

import type { PptxElement } from 'ooxml-core/pptx';

import type { DeviceBox, ExactBoxSource } from './device-pixel-exact';
import { exactDeviceBox } from './device-pixel-exact';

const parents = new WeakMap<object, PptxElement>();

/** Remember `group` as the parent of each of its children (idempotent, cheap). */
export function registerGroupChildren(group: PptxElement): void {
	if (group.type !== 'group') {
		return;
	}
	for (const child of group.children) {
		parents.set(child, group);
	}
}

/** The absolute placement of an element: where it is painted and where PowerPoint paints it. */
export interface SnapPlacement {
	/** Painted position and size in stage pixels. */
	x: number;
	y: number;
	/** The box PowerPoint positions it at, in stage pixels. */
	exact: DeviceBox;
}

const EMU_PER_PX = 9525;

/**
 * The exact origin of a group in its parent's space: its EMU where that still
 * describes the model (a group moved since load keeps the model's value). Not
 * quantised to PowerPoint's master unit: a group child's position comes out of
 * the group transform unrounded (measured at 1280 and 1920 px wide).
 */
function exactOrigin(group: PptxElement): { x: number; y: number } {
	const emuX = (group as ExactBoxSource).xEmu;
	const emuY = (group as ExactBoxSource).yEmu;
	return {
		x:
			typeof emuX === 'number' && Math.round(emuX / EMU_PER_PX) === Math.round(group.x)
				? emuX / EMU_PER_PX
				: group.x,
		y:
			typeof emuY === 'number' && Math.round(emuY / EMU_PER_PX) === Math.round(group.y)
				? emuY / EMU_PER_PX
				: group.y,
	};
}

/**
 * The stage placement of `element`, or `undefined` when an ancestor group is
 * rotated or flipped (no snap). A top-level element places at its own `x` / `y`
 * and its {@link exactDeviceBox}, exactly as before.
 */
export function snapPlacement(element: PptxElement): SnapPlacement | undefined {
	const own = exactDeviceBox(element as ExactBoxSource);
	let group = parents.get(element);
	if (!group) {
		return { x: element.x, y: element.y, exact: own };
	}
	let shownX = 0;
	let shownY = 0;
	let trueX = 0;
	let trueY = 0;
	for (let depth = 0; group && depth < 64; depth++) {
		if (group.rotation || group.flipHorizontal || group.flipVertical) {
			return undefined;
		}
		const origin = exactOrigin(group);
		shownX += group.x;
		shownY += group.y;
		trueX += origin.x;
		trueY += origin.y;
		group = parents.get(group);
	}
	// Local offsets come from the model (they are unrounded and already scaled by
	// the group).
	return {
		// Chromium snaps an inline `<svg>` root (every stroke overlay and connector)
		// to a whole stage pixel, so that is where the painted stroke really starts.
		x: Math.round(shownX + element.x),
		y: Math.round(shownY + element.y),
		exact: {
			x: trueX + element.x,
			y: trueY + element.y,
			width: element.width,
			height: element.height,
		},
	};
}

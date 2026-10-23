import { createGestureController, elementIdSelector } from 'ooxml-ui/pptx';

import { syncSelectionHandleOverlay } from '../../utils/selection-handle-overlay';

/**
 * The rotate knob's drag, on the shared `createGestureController` that every
 * binding's move / resize / rotate runs on. It works in client px (scale 1,
 * origin 0): the element's on-screen box is the gesture's box, so its centre
 * is the pivot whatever the zoom. The rotation is anchored to the press, so an
 * off-centre grab does not jump; Shift snaps to 15 degrees. Frames repaint the
 * element's own transform (and its detached handles) without React state;
 * release commits whole degrees, and only when they changed.
 */
export interface RotateKnobGesture {
	/** The knob button that was pressed; scopes the element lookup to its viewer. */
	button: HTMLElement;
	elementId: string;
	pointer: { clientX: number; clientY: number };
	pointerId?: number;
	/** The element's authored rotation, the drag's baseline. */
	rotation: number;
	/** Element transform sans rotation (flips/skews), painted after the rotation. */
	nonRotationTransform?: string;
	onRotate?: (elementId: string, rotationDeg: number) => void;
}

/** Whole degrees in [0, 360). */
const wholeDegrees = (deg: number): number => Math.round(((deg % 360) + 360) % 360) % 360;

export function startRotateKnobGesture(options: RotateKnobGesture): void {
	const { button, elementId, rotation, onRotate } = options;
	// Resolve the REAL element node by id, scoped to this viewer's
	// `[data-pptx-viewport]`: the handles live in a stage-level overlay that is a
	// sibling of the element, and two viewers of one deck on a page both render
	// the selected id. The `document` fallback only serves bare unit mounts.
	const scope = button.closest('[data-pptx-viewport]') ?? document;
	const wrapper = scope.querySelector<HTMLElement>(elementIdSelector(elementId));
	if (!wrapper) {
		return;
	}
	const rect = wrapper.getBoundingClientRect();
	// Rotation precedes the flip/skew transforms, matching the resting order.
	const base = options.nonRotationTransform ? ` ${options.nonRotationTransform}` : '';
	const paint = (deg: number): void => {
		const transform = `rotate(${deg}deg)${base}`;
		wrapper.style.transform = transform;
		syncSelectionHandleOverlay(wrapper, elementId, { transform });
	};
	if (options.pointerId !== undefined) {
		button.setPointerCapture?.(options.pointerId);
	}
	const gesture = createGestureController({
		getScale: () => 1,
		getElementBox: () => ({
			x: rect.left,
			y: rect.top,
			width: rect.width,
			height: rect.height,
			rotation,
		}),
		getSiblings: () => [],
		getStageOrigin: () => ({ left: 0, top: 0 }),
		onStart: () => undefined,
		onPreview: (transform) => paint(wholeDegrees(transform.rotation)),
		onEnd: (transform) => {
			const deg = transform ? wholeDegrees(transform.rotation) : rotation;
			if (deg !== rotation) {
				onRotate?.(elementId, deg);
			}
		},
	});
	// No pointer id: the knob follows the pointer that pressed it whatever id the
	// compatibility events carry, as it always has.
	gesture.begin('rotate', elementId, {
		clientX: options.pointer.clientX,
		clientY: options.pointer.clientY,
		shiftKey: false,
	});
}

import { DRAG_DEAD_ZONE_PX } from '../render/editor-gestures';
import { resolveTopLevelElementId } from '../render/element-hit-test';

export const ANIMATION_PAINTER_SCOPE = '[data-pptx-theme-root]';

/** Recognize a canvas tap even if the binding replaces the shape node when selecting it. */
export function connectAnimationPainterInput(
	root: HTMLElement,
	actions: { isArmed: () => boolean; apply: (id: string) => void; cancel: () => void },
): () => void {
	let press: { pointerId: number; elementId: string; x: number; y: number } | undefined;
	const view = root.ownerDocument.defaultView;
	const targetId = (target: EventTarget | null): string | null => {
		const viewport = target instanceof Element ? target.closest('[data-pptx-viewport]') : null;
		return viewport?.closest(ANIMATION_PAINTER_SCOPE) === root
			? resolveTopLevelElementId(target, viewport)
			: null;
	};
	const onDown = (event: PointerEvent): void => {
		if (!event.isPrimary || event.button !== 0) {
			return;
		}
		const elementId = actions.isArmed() ? targetId(event.target) : null;
		press = elementId
			? { pointerId: event.pointerId, elementId, x: event.clientX, y: event.clientY }
			: undefined;
	};
	const clearPressOnMove = (event: PointerEvent): void => {
		if (
			press?.pointerId === event.pointerId &&
			(Math.abs(event.clientX - press.x) > DRAG_DEAD_ZONE_PX ||
				Math.abs(event.clientY - press.y) > DRAG_DEAD_ZONE_PX)
		) {
			press = undefined;
		}
	};
	const onUp = (event: PointerEvent): void => {
		clearPressOnMove(event);
		if (press?.pointerId === event.pointerId) {
			const { elementId } = press;
			press = undefined;
			if (actions.isArmed()) {
				actions.apply(elementId);
			}
		}
	};
	const onCancel = (event: PointerEvent): void => {
		if (press?.pointerId === event.pointerId) {
			press = undefined;
		}
	};
	const onClick = (event: MouseEvent): void => {
		if (event.detail === 0 && actions.isArmed()) {
			const elementId = targetId(event.target);
			if (elementId) {
				actions.apply(elementId);
			}
		}
	};
	const onEscape = (event: KeyboardEvent): void => {
		if (event.key === 'Escape' && actions.isArmed()) {
			event.preventDefault();
			event.stopPropagation();
			press = undefined;
			actions.cancel();
		}
	};
	root.addEventListener('pointerdown', onDown, true);
	view?.addEventListener('pointermove', clearPressOnMove, true);
	view?.addEventListener('pointerup', onUp, true);
	view?.addEventListener('pointercancel', onCancel, true);
	root.addEventListener('click', onClick, true);
	root.addEventListener('keydown', onEscape, true);
	return () => {
		root.removeEventListener('pointerdown', onDown, true);
		view?.removeEventListener('pointermove', clearPressOnMove, true);
		view?.removeEventListener('pointerup', onUp, true);
		view?.removeEventListener('pointercancel', onCancel, true);
		root.removeEventListener('click', onClick, true);
		root.removeEventListener('keydown', onEscape, true);
	};
}

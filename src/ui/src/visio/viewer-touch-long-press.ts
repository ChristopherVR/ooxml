import type { ViewerController } from './controller';
import type { Gesture } from './viewer-gesture-state';
import { LONG_PRESS_MS } from './viewer-touch-math';

/**
 * Holding a finger still on a shape, then lifting it, opens its context menu like a right-click.
 * The menu opens on release, as Windows does: a popover opened under a finger that is still down
 * is light-dismissed by that finger's own pointerup.
 */
export function armLongPress(
	view: (Window & typeof globalThis) | null,
	drag: Gesture,
	shape: Element,
	event: PointerEvent,
	live: () => boolean,
): void {
	if (!view) return;
	const { clientX, clientY } = event;
	drag.longPress = view.setTimeout(() => {
		drag.longPress = undefined;
		if (drag.started || !live()) return;
		drag.held = () =>
			shape.dispatchEvent(
				new view.MouseEvent('contextmenu', { bubbles: true, cancelable: true, clientX, clientY }),
			);
	}, LONG_PRESS_MS);
}
/**
 * What lifting a finger that never dragged does, if anything: open the context menu after a long
 * press, or select the thin shape the tap landed beside (the browser click hit empty paper).
 */
export function touchRelease(
	drag: Gesture,
	controller: ViewerController,
): ((view: (Window & typeof globalThis) | null) => void) | undefined {
	const held = drag.held;
	if (held) return (view) => void view?.setTimeout(held, 100);
	const target = drag.target;
	if (drag.touch && drag.nearby && target) return () => controller.selectShapes([target]);
	return undefined;
}

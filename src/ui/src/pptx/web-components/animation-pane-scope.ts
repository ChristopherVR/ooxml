import { createAnimationPaneController } from '../render/controller/animation-pane-controller';
import { ANIMATION_PAINTER_SCOPE } from './animation-painter-input';

const controllers = new WeakMap<Element, ReturnType<typeof createAnimationPaneController>>();

/** Store adapter shared by the native controls mounted by each binding. */
export function animationPaneFor(
	element: HTMLElement,
): ReturnType<typeof createAnimationPaneController> | undefined {
	const root = element.closest(ANIMATION_PAINTER_SCOPE);
	if (!root) return undefined;
	let controller = controllers.get(root);
	if (!controller) {
		controller = createAnimationPaneController();
		controllers.set(root, controller);
	}
	return controller;
}

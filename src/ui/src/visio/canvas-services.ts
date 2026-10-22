import type { ViewerController } from './controller';
import type { ClipboardOperation } from './viewer-clipboard';

/**
 * What a mounted standalone canvas can do beyond the controller: the browser clipboard transport
 * and the drawing surface. Toolbar props and the shapes pane look these up by controller, so a
 * custom UI never reaches into the canvas element.
 */
export interface CanvasServices {
	readonly viewport: HTMLElement;
	/** Run a clipboard operation; rejects with the reason when it is unavailable. */
	clipboard(operation: ClipboardOperation): Promise<void>;
	announce(message: string): void;
}
const mounted = new WeakMap<ViewerController, CanvasServices>();
const watchers = new WeakMap<
	ViewerController,
	Set<(services: CanvasServices | undefined) => void>
>();
const tell = (controller: ViewerController) => {
	for (const watcher of [...(watchers.get(controller) ?? [])]) watcher(mounted.get(controller));
};
export function registerCanvasServices(
	controller: ViewerController,
	services: CanvasServices,
): () => void {
	mounted.set(controller, services);
	tell(controller);
	return () => {
		if (mounted.get(controller) !== services) return;
		mounted.delete(controller);
		tell(controller);
	};
}
/** Follow the canvas mounted on a controller: called now, and whenever one mounts or unmounts. */
export function watchCanvasServices(
	controller: ViewerController,
	watcher: (services: CanvasServices | undefined) => void,
): () => void {
	const set = watchers.get(controller) ?? new Set();
	watchers.set(controller, set);
	set.add(watcher);
	watcher(mounted.get(controller));
	return () => set.delete(watcher);
}
/** The services of the canvas mounted on this controller, if any. */
export function canvasServices(controller: ViewerController): CanvasServices | undefined {
	return mounted.get(controller);
}

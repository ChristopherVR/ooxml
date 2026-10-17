import { subscribeBarFacePicturePixelSamples } from '../render/chart-bar3d-face-picture-sample';

interface ChartSubscription {
	container: HTMLElement;
	unsubscribe: () => void;
}
const subscriptions = new WeakMap<Element, ChartSubscription>();

/** A chart's global sampler subscription must end when its DOM owner is retired. */
export function subscribeChartPixelRepaint(
	container: HTMLElement,
	repaint: (container: HTMLElement) => void,
): void {
	const entry: ChartSubscription = { container, unsubscribe: () => {} };
	entry.unsubscribe = subscribeBarFacePicturePixelSamples(() => {
		if (!entry.container.isConnected) {
			disposeChartPixelSubscriptions(entry.container);
			return;
		}
		repaint(entry.container);
	});
	subscriptions.set(container, entry);
}

/** Retire subscriptions on a chart root or all chart descendants of a stage/host. */
export function disposeChartPixelSubscriptions(root: Element): void {
	const dispose = (node: Element): void => {
		subscriptions.get(node)?.unsubscribe();
		subscriptions.delete(node);
	};
	dispose(root);
	for (const node of root.querySelectorAll('.pptxv-chart')) dispose(node);
}

/** Presentation builds replace children while retaining the animated wrapper. */
export function adoptChartPixelSubscription(from: Element, to: HTMLElement): void {
	disposeChartPixelSubscriptions(to);
	const entry = subscriptions.get(from);
	if (entry) {
		subscriptions.delete(from);
		entry.container = to;
		subscriptions.set(to, entry);
	}
}

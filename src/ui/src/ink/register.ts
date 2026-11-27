import { defineColorGrid } from '../form/color-grid';

/**
 * Defines the Draw tab's elements (`office-ui-ink-pens`, `office-ui-ink-canvas`), loaded on first
 * use: Word and Excel call it when their Draw tab is first shown, so an editor that never draws
 * never downloads them. Idempotent; a no-op without a DOM.
 */
export async function registerInkElements(registry?: CustomElementRegistry): Promise<void> {
	const [{ defineInkPens }, { defineInkCanvas }] = await Promise.all([
		import('./ink-pens'),
		import('./ink-canvas'),
	]);
	// The pen menu's colours.
	defineColorGrid(registry);
	defineInkPens(registry);
	defineInkCanvas(registry);
}

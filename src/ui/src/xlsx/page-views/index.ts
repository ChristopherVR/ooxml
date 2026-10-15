// Workbook views other than Normal. The active sheet's `view.mode` (`sheetView/@view`) decides:
// Page Layout swaps the grid for the page view; Page Break Preview stays on the grid, which paints
// the break lines itself (grid/page-break-layer.ts). This module only follows the model.
import type { EditorContext } from 'ooxml-core/xlsx/ui';
import { createPageLayoutView } from './page-layout-view';

export interface PageViews {
	/** Re-reads the mode and the zoom (the status bar slider changes the latter). */
	refresh(): void;
	destroy(): void;
}

export function mountPageViews(
	ctx: EditorContext,
	host: HTMLElement,
	fileName: () => string,
): PageViews {
	const view = createPageLayoutView(ctx, host, fileName);
	let shown: string | undefined;
	const sync = () => {
		const mode = ctx.workbook()?.sheets[ctx.activeSheet()]?.view?.mode;
		const next = mode ?? 'normal';
		if (next !== shown) {
			const wasLayout = shown === 'pageLayout';
			shown = next;
			host.dataset.view = next;
			view.setActive(next === 'pageLayout');
			if (wasLayout && next !== 'pageLayout') ctx.grid()?.focus();
		}
	};
	const off = ctx.onModelChange(sync);
	sync();
	return {
		refresh() {
			sync();
			view.refresh();
		},
		destroy() {
			off();
			view.destroy();
		},
	};
}

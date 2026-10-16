import { resetChartFormatTarget } from 'ooxml-ui/pptx';
import type { PptxChartData } from 'pptx-viewer-core';

import type { Store, ViewerState } from '../state';
import type { ApplyToSelected } from './editor-apply-to-selected';

/**
 * The chart Format task pane and Chart Format > Current Selection: which chart
 * element they target, whether the pane is open, and the two edits they make
 * (a pane field commit and Reset to Match Style). Decisions (the element list,
 * the field edits, the reset) are the shared `render/chart-format` module; this
 * file only owns the store and the history-integrated patch.
 */
export interface ChartFormatActions {
	/** Make `target` the current chart element (the pane follows). */
	selectChartFormatTarget(target: string): void;
	/** Open the pane, on `target` when given, else on the current element. */
	openChartFormatPane(target?: string): void;
	closeChartFormatPane(): void;
	/** A committed pane edit: the whole new chart model of the selected chart. */
	applyChartFormatData(chartData: PptxChartData): void;
	/** Reset to Match Style on the current chart element. */
	resetChartFormatSelection(): void;
}

/** The current chart element of the selected chart (the chart area by default). */
export function chartFormatTargetOf(state: ViewerState): string {
	return state.chartFormat && state.chartFormat.elementId === state.selectedElementId
		? state.chartFormat.target
		: 'chartArea';
}

export function createChartFormatActions(deps: {
	store: Store<ViewerState>;
	applyToSelected: ApplyToSelected;
}): ChartFormatActions {
	const { store, applyToSelected } = deps;
	const select = (target: string): void => {
		const id = store.get().selectedElementId;
		if (id) {
			store.set({ chartFormat: { elementId: id, target } });
		}
	};
	return {
		selectChartFormatTarget: select,
		openChartFormatPane(target) {
			if (target) {
				select(target);
			}
			store.set({ chartFormatPaneOpen: true });
		},
		closeChartFormatPane() {
			store.set({ chartFormatPaneOpen: false });
		},
		applyChartFormatData(chartData) {
			applyToSelected((el) => (el.type === 'chart' ? { chartData } : {}));
		},
		resetChartFormatSelection() {
			const target = chartFormatTargetOf(store.get());
			applyToSelected((el) =>
				el.type === 'chart' && el.chartData
					? { chartData: resetChartFormatTarget(el.chartData, target) }
					: {},
			);
		},
	};
}

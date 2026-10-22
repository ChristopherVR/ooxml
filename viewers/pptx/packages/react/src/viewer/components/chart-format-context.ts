/**
 * The chart Format pane's link to the viewer: the selected chart, the chart
 * element it targets (`chartArea`, `series:0`, `axis:1`, ...), whether the
 * pane is open, and the commands every surface shares (Chart Format >
 * Current Selection, the chart's click/double-click, the context menu and the
 * pane itself).
 *
 * Provided once by `PowerPointViewer`, like `ShapeFormatContext`: the
 * surfaces sit in different subtrees. Absent a provider (a surface rendered
 * standalone in a unit test) every consumer treats the commands as
 * unavailable.
 */
import type { PptxChartData, PptxElement } from 'pptx-viewer-core';
import { createContext, useContext } from 'react';

export interface ChartFormatCommands {
	/** The selected chart element, or null when the selection is not one chart. */
	chart: PptxElement | null;
	/** The deck is editable and the viewer is in edit or master mode. */
	editable: boolean;
	/** The current chart element (`chartArea`, `series:0`, ...). */
	target: string;
	/** The Format pane is open. */
	paneOpen: boolean;
	/** Make `target` current; `chartId` names the chart when it is not selected yet. */
	setTarget: (target: string, chartId?: string) => void;
	/** Open the pane, optionally on another element first. */
	openPane: (target?: string) => void;
	closePane: () => void;
	/** Apply a whole new chart model to the selected chart (update + history). */
	applyChartData: (chartData: PptxChartData) => void;
	/** Reset to Match Style on the current element. */
	reset: () => void;
}

export const ChartFormatContext = /* @__PURE__ */ createContext<ChartFormatCommands | null>(null);

export function useChartFormatContext(): ChartFormatCommands | null {
	return useContext(ChartFormatContext);
}

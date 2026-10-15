import { chartSelectionHomeControls, dispatchChartSelectionRequest } from 'ooxml-ui/pptx';
import type { PptxElement } from 'pptx-viewer-core';

import type { EditActions } from '../../../editor/editor-edit-ops';
import type { Translator } from '../../../i18n';
import { createSharedHomeStrip } from '../home/shared-strip';

export interface ChartSelectionState {
	editable: boolean;
	/** The selected element; the strip is live only for a chart. */
	element: PptxElement | undefined;
	/** The current chart element (`chartArea`, `series:0`, ...). */
	chartFormatTarget?: string;
}

export interface ChartSelectionStrip {
	el: HTMLElement;
	update(state: ChartSelectionState): void;
}

type ChartFormatHandlers = Pick<
	EditActions,
	'selectChartFormatTarget' | 'openChartFormatPane' | 'resetChartFormatSelection'
>;

/**
 * Chart Format > Current Selection: the shared
 * `pptx-ui-ribbon-home-chart-selection` strip (element drop-down, Format
 * Selection, Reset to Match Style), routed to the chart format edit actions.
 */
export function createChartSelectionStrip(
	doc: Document,
	t: Translator,
	handlers: ChartFormatHandlers,
): ChartSelectionStrip {
	const strip = createSharedHomeStrip(doc, t, 'chart-selection', (intent) =>
		dispatchChartSelectionRequest(intent, {
			select: (target) => handlers.selectChartFormatTarget(target),
			formatSelection: () => handlers.openChartFormatPane(),
			reset: () => handlers.resetChartFormatSelection(),
		}),
	);
	const update = (state: ChartSelectionState) =>
		strip.set(
			chartSelectionHomeControls({
				chartData: state.element?.type === 'chart' ? state.element.chartData : undefined,
				targetKey: state.chartFormatTarget,
				editable: state.editable,
				translate: t,
			}),
		);
	update({ editable: false, element: undefined });
	return { el: strip.el, update };
}

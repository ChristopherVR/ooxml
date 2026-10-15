/**
 * Chart Format > Current Selection: the element drop-down, Format Selection
 * and Reset to Match Style, drawn by the shared
 * `pptx-ui-ribbon-home-chart-selection` strip. Bindings build the state with
 * {@link chartSelectionHomeControls} and route its requests through
 * {@link dispatchChartSelectionRequest}.
 *
 * @module render/chart-format/chart-selection-strip
 */
import type { Translate } from 'ooxml-core/i18n';
import type { PptxChartData, PptxChartFormatTargetKey } from 'ooxml-core/pptx';

import { column, control, text } from '../ribbon-home-family-helpers';
import type {
	RibbonHomeFamilySpec,
	RibbonHomeIntent,
	RibbonHomeViewState,
} from '../ribbon-home-spec';
import {
	listChartFormatTargets,
	resolveChartFormatTarget,
	translateChartFormatText,
} from './chart-format-targets';

export const CHART_SELECTION_ELEMENTS = 'chartFormat.currentSelection.chartElements';
export const CHART_SELECTION_FORMAT = 'chartFormat.currentSelection.formatSelection';
export const CHART_SELECTION_RESET = 'chartFormat.currentSelection.resetToMatchStyle';

/** The strip's controls: PowerPoint's three-row Current Selection group. */
export const CHART_SELECTION_FAMILY: RibbonHomeFamilySpec = {
	clusters: [
		column(
			control(
				CHART_SELECTION_ELEMENTS,
				'pptx.chartFormat.chartElements',
				'Chart Elements',
				undefined,
				{
					kind: 'select',
					icon: false,
					attrs: { 'data-testid': 'chart-format-elements' },
				},
			),
			control(
				CHART_SELECTION_FORMAT,
				'pptx.chartFormat.formatSelection',
				'Format Selection',
				'chart-format-selection',
				{ text: text('pptx.chartFormat.formatSelection', 'Format Selection') },
			),
			control(
				CHART_SELECTION_RESET,
				'pptx.chartFormat.resetToMatchStyle',
				'Reset to Match Style',
				'chart-format-reset',
				{ text: text('pptx.chartFormat.resetToMatchStyle', 'Reset to Match Style') },
			),
		),
	],
};

export interface ChartSelectionInput {
	/** The selected chart's model, or undefined when no chart is selected. */
	chartData: PptxChartData | undefined;
	/** The current chart element (`chartArea`, `series:0`, ...). */
	targetKey: PptxChartFormatTargetKey | undefined;
	editable: boolean;
	translate: Translate;
}

/** The strip's state: the element rows (translated), the current one and gating. */
export function chartSelectionHomeControls(
	input: ChartSelectionInput,
): RibbonHomeViewState['controls'] {
	const disabled = !input.editable || !input.chartData;
	const t = (key: string, params?: Readonly<Record<string, string | number>>) =>
		input.translate(key, params as Record<string, string>);
	const items = input.chartData
		? listChartFormatTargets(input.chartData).map((option) => ({
				value: option.key,
				label: translateChartFormatText(option.label, t),
			}))
		: [];
	const value = resolveChartFormatTarget(input.chartData, input.targetKey).key;
	return {
		[CHART_SELECTION_ELEMENTS]: { disabled: !input.chartData, items, value },
		[CHART_SELECTION_FORMAT]: { disabled: !input.chartData },
		[CHART_SELECTION_RESET]: { disabled },
	};
}

/** What a binding does for each Current Selection request. */
export interface ChartSelectionHandlers {
	/** The drop-down picked an element: make it the current one (the pane follows). */
	select(targetKey: PptxChartFormatTargetKey): void;
	/** Format Selection: open the Format pane on the current element. */
	formatSelection(): void;
	/** Reset to Match Style on the current element. */
	reset(): void;
}

/** Route a `home-request` from the Current Selection strip. */
export function dispatchChartSelectionRequest(
	intent: RibbonHomeIntent,
	handlers: ChartSelectionHandlers,
): void {
	if (intent.id === CHART_SELECTION_ELEMENTS && typeof intent.value === 'string') {
		handlers.select(intent.value);
	} else if (intent.id === CHART_SELECTION_FORMAT) {
		handlers.formatSelection();
	} else if (intent.id === CHART_SELECTION_RESET) {
		handlers.reset();
	}
}

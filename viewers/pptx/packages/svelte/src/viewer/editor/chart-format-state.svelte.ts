import { getContext, setContext } from 'svelte';

/**
 * The chart Format pane's state: which chart element it targets (Chart
 * Format > Current Selection, a click or double-click on the chart, the
 * context menu) and whether the pane is open. Every decision about what a
 * click hits and what an edit writes lives in `ooxml-ui/pptx`; this is only
 * the runes holder the ribbon strip, the canvas chart and the pane share.
 *
 * The target belongs to one chart element: selecting another chart reads as
 * the chart area again, as in PowerPoint, without an effect to reset it.
 */
export class ChartFormatState {
	#owner = $state<string | null>(null);
	#target = $state('chartArea');
	/** The pane was opened (and not closed since). */
	open = $state(false);

	/** The current target of chart `elementId`. */
	targetFor(elementId: string | null | undefined): string {
		return elementId && elementId === this.#owner ? this.#target : 'chartArea';
	}

	/** Make `target` the current element of chart `elementId`. */
	setTarget(elementId: string, target: string): void {
		this.#owner = elementId;
		this.#target = target || 'chartArea';
	}

	/** Open the pane on chart `elementId`, optionally retargeting it first. */
	openPane(elementId: string, target?: string): void {
		if (target !== undefined) {
			this.setTarget(elementId, target);
		} else if (this.#owner !== elementId) {
			this.setTarget(elementId, 'chartArea');
		}
		this.open = true;
	}

	close(): void {
		this.open = false;
	}
}

const ChartFormatContextKey = Symbol('pptx-svelte-chart-format');

/** Provide the viewer's chart format state to the canvas charts (root only). */
export function provideChartFormat(state: ChartFormatState): void {
	setContext(ChartFormatContextKey, state);
}

/** The viewer's chart format state, or undefined outside a viewer (thumbnails, tests). */
export function useChartFormat(): ChartFormatState | undefined {
	return getContext<ChartFormatState | undefined>(ChartFormatContextKey);
}

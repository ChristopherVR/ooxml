import { resetChartFormatTarget } from 'ooxml-ui/pptx';
import type { ChartPptxElement, PptxChartData, PptxElement } from 'pptx-viewer-core';
import type { ComputedRef, InjectionKey, Ref } from 'vue';
import { computed, inject, provide, ref, watch } from 'vue';

/**
 * chart-format-pane (Vue): state of PowerPoint's chart Format task pane and of
 * Chart Format > Current Selection, shared through provide/inject by the pane
 * host, the ribbon strip, the chart canvas handlers and the context menu (the
 * same reason `MergeCropKey` is injected: none of them sit on one prop path).
 *
 * The decisions (which element a click hits, what a field writes, what the
 * strip offers) all live in `ooxml-ui/pptx`; this only holds the current chart
 * element key, the open flag and the history-tracked commit.
 */
export interface ChartFormatController {
	/** The single selected chart, or null. */
	chart: ComputedRef<ChartPptxElement | null>;
	/** Editable deck and a chart selected. */
	editable: ComputedRef<boolean>;
	/** The current chart element (`chartArea`, `series:0`, `axis:1`, ...). */
	target: Ref<string>;
	/** Whether the Format pane is open. */
	open: Ref<boolean>;
	setTarget: (target: string) => void;
	/** Open the pane, optionally on `target`. */
	openPane: (target?: string) => void;
	close: () => void;
	/** Commit a whole new chart model for the selected chart (one undo step). */
	applyChartData: (chartData: PptxChartData) => void;
	/** Reset to Match Style on the current element. */
	reset: () => void;
}

export const ChartFormatKey: InjectionKey<ChartFormatController> = Symbol('pptx-vue-chart-format');

export interface UseChartFormatPaneInput {
	/** The primary selected element. */
	selectedElement: () => PptxElement | null | undefined;
	/** Edit mode, editable deck, not presenting. */
	canEdit: () => boolean;
	/** History-tracked element update (`ops.updateElement`). */
	updateElement: (id: string, patch: Partial<PptxElement>) => void;
}

/** Create the controller and provide it to the viewer's descendants. */
export function useChartFormatPane(input: UseChartFormatPaneInput): ChartFormatController {
	const chart = computed<ChartPptxElement | null>(() => {
		const el = input.selectedElement();
		return el?.type === 'chart' && el.chartData ? (el as ChartPptxElement) : null;
	});
	const editable = computed(() => Boolean(chart.value) && input.canEdit());
	const target = ref('chartArea');
	const open = ref(false);
	// A different chart (or none) starts again at its chart area; no chart closes the pane.
	watch(
		() => chart.value?.id,
		(id, previous) => {
			if (id !== previous) {
				target.value = 'chartArea';
			}
			if (!id) {
				open.value = false;
			}
		},
	);
	function applyChartData(chartData: PptxChartData): void {
		const el = chart.value;
		if (el && input.canEdit()) {
			input.updateElement(el.id, { chartData } as Partial<PptxElement>);
		}
	}
	const controller: ChartFormatController = {
		chart,
		editable,
		target,
		open,
		setTarget: (next) => {
			target.value = next;
		},
		openPane: (next) => {
			if (next) {
				target.value = next;
			}
			if (chart.value) {
				open.value = true;
			}
		},
		close: () => {
			open.value = false;
		},
		applyChartData,
		reset: () => {
			const data = chart.value?.chartData;
			if (data) {
				applyChartData(resetChartFormatTarget(data, target.value));
			}
		},
	};
	provide(ChartFormatKey, controller);
	return controller;
}

/** The injected controller, or undefined outside a viewer (isolated tests, thumbnails). */
export function injectChartFormat(): ChartFormatController | undefined {
	return inject(ChartFormatKey, undefined);
}

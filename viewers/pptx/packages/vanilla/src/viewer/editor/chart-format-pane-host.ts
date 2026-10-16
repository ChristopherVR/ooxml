import { registerPptxWebControls } from 'ooxml-ui/pptx';
import type {
	ChartFormatChangeEvent,
	ChartFormatTargetEvent,
	PptxUiChartFormatPaneElement,
} from 'ooxml-ui/pptx';
import type { PptxElement } from 'pptx-viewer-core';

import type { Translator } from '../i18n';
import { CHART_FORMAT_REQUEST_EVENT } from '../render/elements/chart-editable';
import type { ChartFormatRequestEvent } from '../render/elements/chart-editable';
import type { Store, ViewerState } from '../state';
import type { ChartFormatActions } from './editor-chart-format-actions';
import { chartFormatTargetOf } from './editor-chart-format-actions';

export interface ChartFormatPaneHostDeps {
	doc: Document;
	store: Store<ViewerState>;
	getTranslator(): Translator;
	actions: ChartFormatActions;
	selectedElement(state: ViewerState): PptxElement | undefined;
}

export interface ChartFormatPaneHost {
	/** Listen for chart presses under `root` and dock the pane inside its editor body. */
	attach(root: HTMLElement): void;
	detach(): void;
	/** Show, update or remove the pane from the store. */
	sync(): void;
}

/**
 * Hosts the shared `pptx-ui-chart-format-pane` for the vanilla viewer: shown
 * while the pane is open and a chart is selected in an editable view, fed the
 * selected chart's model and current element, and routed back to the chart
 * format edit actions (one history step per committed field). A press on the
 * selected chart (`CHART_FORMAT_REQUEST_EVENT` from `chart-editable`) retargets
 * it; a double-click opens it.
 */
export function createChartFormatPaneHost(deps: ChartFormatPaneHostDeps): ChartFormatPaneHost {
	const { doc, store, actions } = deps;
	let root: HTMLElement | null = null;
	let pane: PptxUiChartFormatPaneElement | null = null;
	let closing = false;

	const onRequest = (event: Event): void => {
		const { elementId, target, open } = (event as ChartFormatRequestEvent).detail;
		if (store.get().selectedElementId !== elementId) {
			return;
		}
		if (open) {
			actions.openChartFormatPane(target);
		} else {
			actions.selectChartFormatTarget(target);
		}
	};

	const createPane = (): PptxUiChartFormatPaneElement => {
		registerPptxWebControls();
		const node = doc.createElement('pptx-ui-chart-format-pane');
		node.addEventListener('chart-format-change', (event) =>
			actions.applyChartFormatData((event as ChartFormatChangeEvent).detail.chartData),
		);
		node.addEventListener('chart-format-target', (event) =>
			actions.selectChartFormatTarget((event as ChartFormatTargetEvent).detail.target),
		);
		node.addEventListener('chart-format-close', () => actions.closeChartFormatPane());
		return node;
	};

	const remove = (): void => {
		pane?.remove();
		pane = null;
	};

	const sync = (): void => {
		const state = store.get();
		const element = deps.selectedElement(state);
		const chart = element?.type === 'chart' && element.chartData ? element : undefined;
		if (state.chartFormatPaneOpen && !chart && !closing) {
			// The selection moved off the chart: the pane closes, as in PowerPoint.
			// Deferred, because this runs inside a store notification.
			closing = true;
			queueMicrotask(() => {
				closing = false;
				const current = store.get();
				if (current.chartFormatPaneOpen && deps.selectedElement(current)?.type !== 'chart') {
					actions.closeChartFormatPane();
				}
			});
		}
		if (!root || !chart || !state.chartFormatPaneOpen || state.presenting) {
			remove();
			return;
		}
		if (!pane) {
			pane = createPane();
		}
		if (!pane.isConnected) {
			(root.querySelector<HTMLElement>('[data-pptx-chrome="body"]') ?? root).append(pane);
		}
		pane.translateLabel = deps.getTranslator();
		pane.disabled = !state.editable;
		if (pane.chartData !== chart.chartData) {
			pane.chartData = chart.chartData;
		}
		const target = chartFormatTargetOf(state);
		if (pane.target !== target) {
			pane.target = target;
		}
	};

	return {
		attach(next) {
			root?.removeEventListener(CHART_FORMAT_REQUEST_EVENT, onRequest);
			root = next;
			root.addEventListener(CHART_FORMAT_REQUEST_EVENT, onRequest);
			sync();
		},
		detach() {
			root?.removeEventListener(CHART_FORMAT_REQUEST_EVENT, onRequest);
			root = null;
			remove();
		},
		sync,
	};
}

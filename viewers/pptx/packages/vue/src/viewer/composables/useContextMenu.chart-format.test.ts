// oxlint-disable react-hooks/rules-of-hooks
import { mount } from '@vue/test-utils';
import type { PptxElement } from 'pptx-viewer-core';
import { afterEach, describe, expect, it } from 'vitest';
import { computed, defineComponent, h, ref } from 'vue';

import { useChartFormatPane } from './chart-format-pane';
import type { ChartFormatController } from './chart-format-pane';
import { useContextMenu } from './useContextMenu';
import type { UseContextMenuResult } from './useContextMenu';
import type { EditorOperations } from './useEditorOperations';

const CHART = {
	id: 'chart-1',
	type: 'chart',
	x: 0,
	y: 0,
	width: 400,
	height: 300,
	chartData: {
		chartType: 'bar',
		categories: ['A'],
		series: [{ name: 'East', values: [1] }],
	},
} as PptxElement;

let host: HTMLElement | null = null;
afterEach(() => {
	host?.remove();
	host = null;
});

function setup(): { menu: UseContextMenuResult; chartFormat: ChartFormatController } {
	let menu!: UseContextMenuResult;
	let chartFormat!: ChartFormatController;
	const selectedElementIds = ref<string[]>([CHART.id]);
	mount(
		defineComponent({
			setup() {
				chartFormat = useChartFormatPane({
					selectedElement: () => CHART,
					canEdit: () => true,
					updateElement: () => {},
				});
				menu = useContextMenu({
					canEdit: () => true,
					findActiveElement: (id) => (id === CHART.id ? CHART : undefined),
					tableSelection: ref(null),
					hasClipboard: computed(() => true),
					canGroup: computed(() => false),
					selectionGroupable: computed(() => true),
					editTemplateMode: ref(false),
					selectedElementIds,
					inlineEditingElementId: ref<string | null>(null),
					inspectorOpen: ref(false),
					enterInlineEdit: () => {},
					ops: {} as EditorOperations,
					cutElement: () => {},
					copyElement: () => {},
					pasteElement: () => {},
					onGroup: () => {},
					onUngroup: () => {},
					openHyperlinkDialog: () => {},
					chartFormat,
				});
				return () => h('div');
			},
		}),
	);
	return { menu, chartFormat };
}

/** A right-click on a series mark of a rendered chart. */
function rightClickSeries(menu: UseContextMenuResult): void {
	host = document.createElement('div');
	host.dataset.elementId = CHART.id;
	host.innerHTML =
		'<svg viewBox="0 0 400 300"><rect data-chart-part="dataPoint" data-chart-series="0" data-chart-point="0"></rect></svg>';
	document.body.append(host);
	const mark = host.querySelector('rect')!;
	const event = new MouseEvent('contextmenu', { bubbles: true, clientX: 5, clientY: 5 });
	Object.defineProperty(event, 'target', { value: mark });
	menu.onCanvasContextMenu(event);
}

describe('useContextMenu chart Format entry', () => {
	it('names the entry after the clicked chart element and opens the pane on it', () => {
		const { menu, chartFormat } = setup();
		rightClickSeries(menu);
		const entry = menu.contextItems.value.find((i) => i.id === 'format-chart-element');
		expect(entry?.label).toBe('Format Data Series...');
		expect(menu.contextItems.value.some((i) => i.id === 'format-shape')).toBeFalsy();
		menu.onContextSelect('format-chart-element');
		expect(chartFormat.open.value).toBeTruthy();
		expect(chartFormat.target.value).toBe('series:0');
	});
});

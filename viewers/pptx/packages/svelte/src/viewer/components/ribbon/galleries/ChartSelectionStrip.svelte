<script lang="ts">
	/**
	 * ChartSelectionStrip: Chart Format > Current Selection, the shared
	 * `pptx-ui-ribbon-home-chart-selection` strip (element drop-down, Format
	 * Selection, Reset to Match Style). The rows, gating and request routing
	 * are shared; this file only maps them onto the editor.
	 */
	import {
		chartSelectionHomeControls,
		dispatchChartSelectionRequest,
		homeSnapshotTranslator,
		resetChartFormatTarget,
	} from 'ooxml-ui/pptx';
	import type { RibbonHomeRequestEvent } from 'ooxml-ui/pptx';

	import { useTranslator } from '../../../../i18n/context';
	import type { EditorState } from '../../../editor/editor-state.svelte';

	const { editor }: { editor: EditorState } = $props();
	const t = useTranslator();
	const chart = $derived(
		editor.selectedElement?.type === 'chart' ? editor.selectedElement : undefined,
	);
	const state = $derived({
		controls: chartSelectionHomeControls({
			chartData: chart?.chartData,
			targetKey: editor.chartFormat.targetFor(chart?.id),
			editable: editor.editable,
			translate: t,
		}),
		translate: homeSnapshotTranslator(['chart-selection'], t),
	});

	function request(event: RibbonHomeRequestEvent): void {
		const current = chart;
		if (!current) {
			return;
		}
		dispatchChartSelectionRequest(event.detail, {
			select: (target) => editor.chartFormat.setTarget(current.id, target),
			formatSelection: () => editor.chartFormat.openPane(current.id),
			reset: () => {
				if (current.chartData && editor.editable) {
					editor.applyElementPatch(current.id, {
						chartData: resetChartFormatTarget(
							current.chartData,
							editor.chartFormat.targetFor(current.id),
						),
					});
				}
			},
		});
	}
</script>

<pptx-ui-ribbon-home-chart-selection {state} onhome-request={request}
></pptx-ui-ribbon-home-chart-selection>

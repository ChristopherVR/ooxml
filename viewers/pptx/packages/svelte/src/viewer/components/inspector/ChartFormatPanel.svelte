<script lang="ts">
	/**
	 * ChartFormatPanel: hosts the shared `pptx-ui-chart-format-pane` (Format
	 * Chart Area / Data Series / Axis / ...) while it is open on the selected
	 * chart. The pane owns the view and builds each new chart model; this file
	 * only applies it through the editor's normal patch + history path.
	 */
	import type { ChartFormatChangeEvent, ChartFormatTargetEvent } from 'ooxml-ui/pptx';

	import { useTranslator } from '../../../i18n/context';
	import type { EditorState } from '../../editor/editor-state.svelte';

	const { editor }: { editor: EditorState } = $props();
	const t = useTranslator();
	const chart = $derived(
		editor.selectedElement?.type === 'chart' ? editor.selectedElement : undefined,
	);

	// Selecting anything but a chart closes the pane, as in PowerPoint.
	$effect(() => {
		if (!chart && editor.chartFormat.open) {
			editor.chartFormat.close();
		}
	});

	function change(event: Event): void {
		const current = chart;
		if (current && editor.editable) {
			editor.applyElementPatch(current.id, {
				chartData: (event as ChartFormatChangeEvent).detail.chartData,
			});
		}
	}

	function retarget(event: Event): void {
		if (chart) {
			editor.chartFormat.setTarget(chart.id, (event as ChartFormatTargetEvent).detail.target);
		}
	}
</script>

{#if editor.chartFormat.open && chart?.chartData}
	<pptx-ui-chart-format-pane
		chartData={chart.chartData}
		target={editor.chartFormat.targetFor(chart.id)}
		translateLabel={t}
		disabled={!editor.editable}
		onchart-format-change={change}
		onchart-format-target={retarget}
		onchart-format-close={() => editor.chartFormat.close()}
	></pptx-ui-chart-format-pane>
{/if}

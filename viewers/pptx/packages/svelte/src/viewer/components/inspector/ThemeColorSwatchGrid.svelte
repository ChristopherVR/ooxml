<script lang="ts">
	/**
	 * PowerPoint's "Theme Colors" grid on the shared `pptx-ui-color-grid`: ten columns (Background 1,
	 * Text 1, Background 2, Text 2, Accent 1..6), each a base swatch over five luminance variants,
	 * built from the loaded deck's theme (`EditorState.themeColorMap`) rather than a hard-coded
	 * Office palette.
	 *
	 * Renders nothing when no deck theme is loaded yet, so callers can render this unconditionally
	 * alongside their existing hex and recent-colour controls.
	 */
	import type { PptxThemeColorRef } from 'pptx-viewer-core';
	import type { OfficeColorPickEvent, ThemeColorPickerCommit } from 'ooxml-ui/pptx';
	import { buildThemeColorSwatchGrid, themeColorGridCommit, themeColorGridView } from 'ooxml-ui/pptx';

	import { useTranslator } from '../../../i18n/context';

	const {
		themeColorMap,
		selectedRef,
		selectedHex,
		disabled = false,
		onpick,
	}: {
		themeColorMap: Record<string, string> | undefined;
		selectedRef?: PptxThemeColorRef;
		selectedHex?: string;
		disabled?: boolean;
		onpick: (commit: ThemeColorPickerCommit) => void;
	} = $props();

	const t = useTranslator();

	const columns = $derived(buildThemeColorSwatchGrid(themeColorMap, t));
	const view = $derived(themeColorGridView(columns, selectedRef, selectedHex));
	const heading = $derived(t('pptx.colorPicker.themeColors'));

	function pick(event: OfficeColorPickEvent): void {
		const commit = themeColorGridCommit(columns, event.detail);
		if (commit) {
			onpick(commit);
		}
	}
</script>

<!-- The inspector keeps its own standard and recent rows, so the grid shows only Theme Colors. -->
{#if columns.length > 0}
	<pptx-ui-color-grid
		class="pptx-svelte-theme-swatch-grid"
		palette={view.palette}
		selectedTheme={view.selectedTheme}
		standardColors={[]}
		themeHeading={heading}
		label={heading}
		{disabled}
		onoffice-color-pick={pick}
	></pptx-ui-color-grid>
{/if}

<style>
	.pptx-svelte-theme-swatch-grid {
		margin-top: 4px;
	}
</style>

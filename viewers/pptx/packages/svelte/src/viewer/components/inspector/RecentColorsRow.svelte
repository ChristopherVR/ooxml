<script lang="ts">
	/**
	 * RecentColorsRow: the "Recent colours" swatch strip shown under a colour
	 * input (wave-4 B6), drawn by the shared `pptx-ui-color-grid` with only its
	 * Recent Colors section (`recentColorsGridView`). One swatch per colour,
	 * most-recent-first; clicking a swatch applies it through THIS picker's own
	 * commit handler (`onselect`), so a section with several colour inputs
	 * (fill, line, ...) gives each its own row over the SAME shared list,
	 * mirroring React's per-`ColorPickerRow` approach. Renders nothing while the
	 * list is empty.
	 */
	import type { OfficeColorPickEvent } from 'ooxml-ui/pptx';
	import { recentColorsGridView } from 'ooxml-ui/pptx';

	import { useTranslator } from '../../../i18n/context';

	const {
		colors,
		onselect,
		disabled = false,
	}: {
		colors: readonly string[];
		onselect: (hex: string) => void;
		/** Mirrors the owning picker's disabled state onto every swatch. */
		disabled?: boolean;
	} = $props();

	const t = useTranslator();

	// Reads the language through `t`, so a switch hands the grid freshly translated labels.
	const view = $derived(recentColorsGridView(colors, t));
</script>

{#if colors.length > 0}
	<pptx-ui-color-grid
		class="pptx-svelte-recent-colors"
		data-testid="pptx-color-recent"
		palette={view.palette}
		standardColors={view.standardColors}
		recentColors={view.recentColors}
		recentHeading={view.recentHeading}
		recentLabel={view.recentLabel}
		label={view.label}
		{disabled}
		onoffice-color-pick={(event: OfficeColorPickEvent) => onselect(event.detail.color)}
	></pptx-ui-color-grid>
{/if}

<style>
	.pptx-svelte-recent-colors {
		margin-top: 4px;
	}
</style>

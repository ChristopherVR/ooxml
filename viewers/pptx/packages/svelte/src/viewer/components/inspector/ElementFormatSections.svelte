<script lang="ts">
	/**
	 * ElementFormatSections: the selected element's Properties as Format pane
	 * sections. Each section holds the existing panels `FORMAT_PANEL_SECTIONS`
	 * assigns to it (the same placement as the other bindings); every panel
	 * still edits through `EditorState`, so nothing here owns editing logic.
	 */
	import type { PptxElement } from 'pptx-viewer-core';
	import { hasTextProperties, isImageLikeElement } from 'pptx-viewer-core';
	import { formatSelectionElements } from 'ooxml-ui/pptx';
	import type { FormatPaneController, FormatSectionView } from 'ooxml-ui/pptx';

	import type { EditorState } from '../../editor/editor-state.svelte';
	import type { InspectorDeckActions } from '../../state/inspector-deck';
	import ActionSettingsPanel from './ActionSettingsPanel.svelte';
	import AltTextSection from './AltTextSection.svelte';
	import ChartSection from './ChartSection.svelte';
	import EffectsPanel from './EffectsPanel.svelte';
	import FillStrokeSection from './FillStrokeSection.svelte';
	import FormatSections from './FormatSections.svelte';
	import GroupInfoSection from './GroupInfoSection.svelte';
	import ImageSection from './ImageSection.svelte';
	import MediaSection from './MediaSection.svelte';
	import OlePropertiesSection from './OlePropertiesSection.svelte';
	import PositionSection from './PositionSection.svelte';
	import QuickStylesGallery from './QuickStylesGallery.svelte';
	import ShapeSection from './ShapeSection.svelte';
	import SmartArtSection from './SmartArtSection.svelte';
	import TableDataGrid from './TableDataGrid.svelte';
	import TableSection from './TableSection.svelte';
	import TextSection from './TextSection.svelte';

	const {
		editor,
		el,
		pane,
		deck,
		mediaDataUrls,
	}: {
		editor: EditorState;
		el: PptxElement;
		pane: FormatPaneController;
		deck?: InspectorDeckActions | undefined;
		mediaDataUrls: Map<string, string>;
	} = $props();

	// The whole selection, primary first, for the shared sections and Mixed tags.
	const elements = $derived.by(() => {
		const selected = editor.selectedElements;
		return formatSelectionElements(
			el,
			selected.map((item) => item.id),
			selected,
		);
	});
	// React gates Quick Styles on shape/text: the presets are shape-fill recipes.
	const canQuickStyle = $derived(el.type === 'shape' || el.type === 'text');
</script>

<FormatSections {elements} {pane}>
	{#snippet section(view: FormatSectionView)}
		{#if view.id === 'font' || view.id === 'paragraph'}
			<!-- A table's cell text is formatted from its Table section and the ribbon. -->
			{#if hasTextProperties(el)}<TextSection {editor} {el} part={view.id} />{/if}
		{:else if view.id === 'fill'}
			{#if canQuickStyle}<QuickStylesGallery {editor} {el} />{/if}
			<ShapeSection {editor} {el} part="fill" />
			<FillStrokeSection {editor} {el} part="fill" />
		{:else if view.id === 'outline'}
			<FillStrokeSection {editor} {el} part="outline" />
			{#if view.arrowheads}<ShapeSection {editor} {el} part="arrowheads" />{/if}
		{:else if view.id === 'effects'}
			<EffectsPanel {editor} {el} />
		{:else if view.id === 'picture'}
			{#if isImageLikeElement(el)}
				<ImageSection {editor} {el} />
				<AltTextSection {editor} {el} />
			{/if}
		{:else if view.id === 'table'}
			<TableDataGrid {editor} {el} />
			<TableSection
				{editor}
				{el}
				tableStyleMap={deck?.tableStyleMap}
				onTableStyleMapChange={deck?.updateTableStyleMap}
				onDeleteTableStyle={deck?.deleteTableStyle}
			/>
		{:else if view.id === 'chart'}
			<ChartSection {editor} />
		{:else if view.id === 'smartArt'}
			{#if el.type === 'smartArt'}<SmartArtSection {editor} {el} />{/if}
		{:else if view.id === 'media'}
			<MediaSection {editor} {mediaDataUrls} />
		{:else if view.id === 'object'}
			<OlePropertiesSection {editor} {el} />
		{:else if view.id === 'arrange'}
			<PositionSection {editor} {el} part="lock" />
			<ShapeSection {editor} {el} part="shapeType" />
			{#if el.type === 'group'}<GroupInfoSection {el} />{/if}
		{:else if view.id === 'sizePosition'}
			<PositionSection {editor} {el} part="geometry" />
		{:else if view.id === 'altText'}
			<AltTextSection {editor} {el} />
		{:else if view.id === 'action'}
			<ActionSettingsPanel {editor} {el} />
		{/if}
	{/snippet}
</FormatSections>

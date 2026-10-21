<script lang="ts">
	/**
	 * QuickStylesGallery: the inspector's Quick Styles (PowerPoint's Shape
	 * Styles presets), mirroring React's `inspector/QuickStylesGallery.tsx`.
	 *
	 * The tiles are the shared `pptx-ui-ribbon-gallery` panel (compact) over
	 * `buildQuickStylesGallery`, so the presets, previews, accessible names and
	 * touch sizes match every other binding. A pick applies the preset through
	 * the shared `shapeQuickStylePatch`, which MERGES over the element's current
	 * `shapeStyle` instead of replacing it (a quick style is an overlay in
	 * PowerPoint, not a reset). A read-only viewer shows the tiles disabled.
	 */
	import type { PptxElement } from 'pptx-viewer-core';
	import {
		buildQuickStylesGallery,
		quickStyleShapeStyle,
		shapeQuickStylePatch,
	} from 'ooxml-ui/pptx';
	import type { RibbonGalleryPickEvent } from 'ooxml-ui/pptx';

	import { useTranslator } from '../../../i18n/context';
	import type { EditorState } from '../../editor/editor-state.svelte';
	import { strictTranslator } from '../ribbon/galleries/gallery-labels';

	const { editor, el }: { editor: EditorState; el: PptxElement } = $props();
	const t = useTranslator();
	const canEdit = $derived(editor.editable);
	const gallery = buildQuickStylesGallery();
	// The element translates its tiles outside Svelte's tracking: read the caption here so a
	// language switch hands it a new translator and the tiles are named again.
	const translateLabel = $derived.by(() => {
		void t('pptx.shape.quickStyles');
		return strictTranslator(t);
	});

	function pick(event: Event): void {
		const style = quickStyleShapeStyle((event as RibbonGalleryPickEvent).detail.itemId);
		if (style) {
			editor.applyElementPatch(el.id, shapeQuickStylePatch(el, style));
		}
	}
</script>

<!-- The caption is not a <label>: a click on it would press the first tile. -->
<div class="pptx-svelte-quick-styles">
	<span class="pptx-svelte-quick-styles-label">{t('pptx.shape.quickStyles')}</span>
	<pptx-ui-ribbon-gallery
		mode="panel"
		compact
		descriptor={gallery}
		{translateLabel}
		disabled={!canEdit}
		ongallery-pick={pick}
	></pptx-ui-ribbon-gallery>
</div>

<style>
	.pptx-svelte-quick-styles {
		display: flex;
		flex-direction: column;
		gap: 4px;
		margin-bottom: 10px;
	}

	.pptx-svelte-quick-styles-label {
		color: var(--pptx-muted-foreground, var(--office-muted-foreground, #6b7280));
	}
</style>

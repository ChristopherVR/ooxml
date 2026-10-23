<script lang="ts">
	/**
	 * EffectsPanel: the shape effects section (outer and inner shadow, glow,
	 * soft edge, reflection, blur, top bevel). The fields are the shared
	 * `SHAPE_EFFECT_SECTIONS` descriptors drawn by the generic
	 * `InspectorFields` renderer; each `Partial<ShapeStyle>` patch is merged
	 * into the element with the shared `shapeQuickStylePatch`. Rendered by
	 * `InspectorPanel` alongside `FillStrokeSection` for shape-property
	 * elements only.
	 */
	import { SHAPE_EFFECT_SECTIONS, shapeQuickStylePatch } from 'ooxml-ui/pptx';
	import type { PptxElement, ShapeStyle } from 'pptx-viewer-core';
	import { hasShapeProperties } from 'pptx-viewer-core';

	import type { EditorState } from '../../editor/editor-state.svelte';
	import InspectorFields from './InspectorFields.svelte';

	const { editor, el }: { editor: EditorState; el: PptxElement } = $props();

	const shapeStyle = $derived(hasShapeProperties(el) ? el.shapeStyle : undefined);

	function onpatch(patch: Partial<ShapeStyle>): void {
		editor.patchSelected(shapeQuickStylePatch(el, patch));
	}

	/** A colour drag's live shade: on the slide at once, no undo step. */
	function onpreview(patch: Partial<ShapeStyle>): void {
		editor.previewElement(el.id, shapeQuickStylePatch(el, patch));
	}
</script>

<div class="pptx-svelte-effects">
	<InspectorFields
		sections={SHAPE_EFFECT_SECTIONS}
		target={shapeStyle}
		{onpatch}
		{onpreview}
		oncolorcommit={(hex) => editor.recordRecentColor(hex)}
	/>
</div>

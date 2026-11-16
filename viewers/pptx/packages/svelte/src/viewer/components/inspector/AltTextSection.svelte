<script lang="ts">
	/**
	 * AltTextSection: the Alt Text fields (description and title) of the
	 * selected element, the shared `ELEMENT_ALT_TEXT_SECTIONS` drawn by the
	 * generic `InspectorFields`. The descriptors decide which fields apply to
	 * `el`'s kind (a picture models only the description), so this component
	 * is mounted for pictures and every other kind by `ElementFormatSections`.
	 */
	import type { PptxElement } from 'pptx-viewer-core';
	import { ELEMENT_ALT_TEXT_SECTIONS, getNonVisualDescriptionFields } from 'ooxml-ui/pptx';

	import type { EditorState } from '../../editor/editor-state.svelte';
	import InspectorFields from './InspectorFields.svelte';

	const { editor, el }: { editor: EditorState; el: PptxElement } = $props();
	const fields = $derived(getNonVisualDescriptionFields(el));
</script>

{#if fields.showAltText || fields.showTitle}
	<div data-pptx-inspector-section="alt-text">
		<InspectorFields
			sections={ELEMENT_ALT_TEXT_SECTIONS}
			target={el}
			disabled={!editor.editable}
			onpatch={(patch) => editor.applyElementPatch(el.id, patch)}
		/>
	</div>
{/if}

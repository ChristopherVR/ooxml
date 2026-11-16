<script lang="ts">
	/**
	 * OlePropertiesSection: OLE object summary (type / file name / link status)
	 * plus the Object Name editor, at parity with React's `ElementMiscPanels.tsx`
	 * OlePropertiesPanel.
	 *
	 * A browser cannot run the native application that owns an embedded OLE
	 * object, so the object itself stays read-only. Its Object Name IS editable:
	 * `p:oleObj/@name` (ECMA-376 SS13.3.4) already parses, saves, and syncs via
	 * collaboration, and shared's `getOleDisplayName` / `getOleAriaLabel` already
	 * read it, so this field was the only piece missing to make it a real,
	 * round-tripping edit.
	 */
	import type { OlePptxElement, PptxElement } from 'pptx-viewer-core';
	import { OLE_OBJECT_NAME_SECTIONS, resolveOleSummaryRows } from 'ooxml-ui/pptx';

	import { useTranslator } from '../../../i18n/context';
	import type { EditorState } from '../../editor/editor-state.svelte';
	import InspectorFields from './InspectorFields.svelte';
	import OleEditorDialog from './OleEditorDialog.svelte';

	const { editor, el }: { editor: EditorState; el: PptxElement } = $props();
	const t = useTranslator();
	const canEdit = $derived(editor.editable);

	const ole = $derived(el as OlePptxElement);

	let isEditorOpen = $state(false);
</script>

<div class="pptx-svelte-ole-info">
	{#if canEdit && !ole.isLinked}
		<button type="button" class="pptx-svelte-ole-edit-content" onclick={() => (isEditorOpen = true)}>
			{t('pptx.ole.editContent')}
		</button>
	{/if}
	<InspectorFields
		sections={OLE_OBJECT_NAME_SECTIONS}
		target={el}
		disabled={!canEdit}
		onpatch={(patch) => editor.applyElementPatch(el.id, patch)}
	/>
	{#each resolveOleSummaryRows(el) as row (row.id)}
		<div class="pptx-svelte-ole-row">
			<span>{t(row.labelKey)}</span>
			{#if row.id === 'linkStatus'}
				<span class="pptx-svelte-ole-badge" class:is-linked={ole.isLinked}>{t(row.valueKey ?? '')}</span>
			{:else}
				<span class="pptx-svelte-ole-value" title={row.value}>{row.value}</span>
			{/if}
		</div>
	{/each}
</div>

<OleEditorDialog {editor} el={ole} open={isEditorOpen} onclose={() => (isEditorOpen = false)} />

<style>
	.pptx-svelte-ole-edit-content {
		width: 100%;
		padding: 6px 8px;
		border: 1px solid var(--pptx-border, var(--office-border, #d1d5db));
		border-radius: var(--pptx-radius, var(--office-radius, 4px));
		background: transparent;
		color: inherit;
		cursor: pointer;
		font: inherit;
		font-size: 11px;
	}
	.pptx-svelte-ole-edit-content:hover {
		background: var(--pptx-accent, var(--office-selected, #e5e7eb));
	}
	.pptx-svelte-ole-info {
		display: flex;
		flex-direction: column;
		gap: 6px;
		font-size: 11px;
	}
	.pptx-svelte-ole-row {
		display: flex;
		align-items: center;
		justify-content: space-between;
		gap: 8px;
	}
	.pptx-svelte-ole-row > span:first-child {
		color: var(--pptx-muted-foreground, var(--office-muted-foreground, #6b7280));
	}
	.pptx-svelte-ole-value {
		overflow: hidden;
		text-overflow: ellipsis;
		white-space: nowrap;
	}
	.pptx-svelte-ole-badge {
		display: inline-flex;
		align-items: center;
		border-radius: 999px;
		padding: 1px 8px;
		font-size: 10px;
		font-weight: 500;
		background: rgba(34, 197, 94, 0.2);
		color: #4ade80;
	}
	.pptx-svelte-ole-badge.is-linked {
		background: rgba(59, 130, 246, 0.2);
		color: #60a5fa;
	}
</style>

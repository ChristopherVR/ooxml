<script lang="ts">
	/**
	 * PositionSection: the universal X / Y / W / H / rotation numeric grid,
	 * shown for every selected element regardless of type. The fields are the
	 * shared `ELEMENT_GEOMETRY_SECTIONS` descriptors drawn by `InspectorFields`. Extracted from the
	 * original monolithic InspectorPanel so the panel can grow element-type-aware
	 * sections without blowing the file-size budget.
	 */
	import type { PptxElement } from 'pptx-viewer-core';
	import type { ElementGeometryPatch, ZoomCoverPatchEvent } from 'ooxml-ui/pptx';
	import {
		sizePositionSections,
		elementLockTogglePatch,
		isElementLocked,
	} from 'ooxml-ui/pptx';
	import Lock from '@lucide/svelte/icons/lock';
	import LockOpen from '@lucide/svelte/icons/lock-open';

	import { useTranslator } from '../../../i18n/context';
	import type { EditorState } from '../../editor/editor-state.svelte';
	import InspectorFields from './InspectorFields.svelte';

	const {
		editor,
		el,
		part,
	}: {
		editor: EditorState;
		el: PptxElement;
		/** `lock` (Arrange) or `geometry` (Size and position); omitted, both. */
		part?: 'lock' | 'geometry';
	} = $props();
	const t = useTranslator();
	const zoomCover = $derived({ element: el, disabled: !editor.editable, translate: t });
	function applyCover(event: ZoomCoverPatchEvent): void {
		commit(event.detail);
	}

	function commit(patch: ElementGeometryPatch): void {
		editor.applyElementPatch(el.id, patch as Partial<PptxElement>);
	}

	// Shared decides both what reads as "locked" and what the toggle writes, so
	// the button's state can never drift from what the canvas enforces.
	const isLocked = $derived(isElementLocked(el));

	function toggleLock(): void {
		editor.applyElementPatch(el.id, { locks: elementLockTogglePatch(!isLocked) } as Partial<PptxElement>);
	}
</script>

{#if part !== 'geometry'}
<div
	class="pptx-svelte-inspector-lock-row"
	data-pptx-inspector-section={part === undefined ? 'transform' : undefined}
>
	<span>{t('pptx.inspector.element')}</span>
	<button
		type="button"
		class="pptx-svelte-inspector-lock-btn"
		onclick={toggleLock}
		title={isLocked ? t('pptx.inspector.unlock') : t('pptx.inspector.lock')}
		aria-pressed={isLocked}
	>
		{#if isLocked}
			<Lock size={14} color="var(--pptx-amber, #f59e0b)" aria-hidden="true" />
		{:else}
			<LockOpen size={14} aria-hidden="true" />
		{/if}
	</button>
</div>
{/if}
{#if part !== 'lock'}
	<div data-pptx-inspector-section={part === 'geometry' ? 'transform' : undefined}>
		<InspectorFields sections={sizePositionSections(el)} target={el} onpatch={commit} />
		{#if el.type === 'zoom'}
			<pptx-ui-zoom-cover state={zoomCover} onzoom-cover-patch={applyCover}></pptx-ui-zoom-cover>
		{/if}
	</div>
{/if}

<style>
	.pptx-svelte-inspector-lock-row {
		display: flex;
		align-items: center;
		justify-content: space-between;
		margin-bottom: 6px;
		color: var(--pptx-muted-foreground, var(--office-muted-foreground, #6b7280));
	}

	.pptx-svelte-inspector-lock-btn {
		display: inline-flex;
		align-items: center;
		justify-content: center;
		padding: 4px;
		border: none;
		border-radius: var(--pptx-radius, var(--office-radius, 4px));
		background: transparent;
		color: inherit;
		cursor: pointer;
	}

	.pptx-svelte-inspector-lock-btn:hover {
		background: var(--pptx-accent, var(--office-selected, #e5e7eb));
	}
</style>

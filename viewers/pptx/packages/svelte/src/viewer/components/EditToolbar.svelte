<script lang="ts">
	/**
	 * EditToolbar: the secondary formatting/editing row shown under the main
	 * viewer toolbar when editing is active. Thin composition only: it lays out
	 * the Insert, text-format, shape-format, and Arrange (z-order) groups, each
	 * of which reads/writes the shared history-tracked {@link EditorState}. All
	 * behaviour lives in those child components + the editor modules.
	 */
	import { useTranslator } from '../../i18n/context';
	import type { EditorState } from '../editor/editor-state.svelte';
	import ArrangeGroup from './ArrangeGroup.svelte';
	import InsertMenu from './InsertMenu.svelte';
	import ShapeFormatGroup from './ShapeFormatGroup.svelte';
	import TextFormatGroup from './TextFormatGroup.svelte';

	const { editor }: { editor: EditorState } = $props();
	const t = useTranslator();
</script>

<div class="pptx-svelte-edittoolbar" role="toolbar" aria-label={t('pptx.inspector.elementProperties')}>
	<InsertMenu {editor} />
	<span class="pptx-svelte-edittoolbar-sep" aria-hidden="true"></span>
	<TextFormatGroup {editor} />
	<span class="pptx-svelte-edittoolbar-sep" aria-hidden="true"></span>
	<ShapeFormatGroup {editor} />
	<span class="pptx-svelte-edittoolbar-sep" aria-hidden="true"></span>
	<ArrangeGroup {editor} />
</div>

<style>
	.pptx-svelte-edittoolbar {
		display: flex;
		align-items: center;
		flex-wrap: wrap;
		gap: 6px;
		padding: 5px 10px;
		background: var(--pptx-card, var(--office-surface, #f3f4f6));
		color: var(--pptx-card-foreground, var(--office-foreground, #1f2937));
		border-bottom: 1px solid var(--pptx-border, var(--office-border, #d1d5db));
		font-family: system-ui, sans-serif;
		font-size: 13px;
		flex: none;
	}

	.pptx-svelte-edittoolbar-sep {
		width: 1px;
		height: 22px;
		background: var(--pptx-border, var(--office-border, #d1d5db));
	}
</style>

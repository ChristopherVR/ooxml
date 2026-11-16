<script lang="ts">
	/**
	 * SlideBackgroundSection: the active slide's own background (the shared
	 * `SLIDE_BACKGROUND_SECTIONS`: solid colour and Hide background graphics,
	 * drawn by `InspectorFields`, plus the Clear Background action), at parity with React's
	 * `SlideBackgroundPanel`, Vue's and Angular's slide-background card. Clear
	 * visibility and enablement come from the shared
	 * `slideBackgroundClearState`, and one click is one undo step through
	 * `EditorState.backgroundOps`.
	 */
	import { SLIDE_BACKGROUND_SECTIONS, slideBackgroundClearState } from 'ooxml-ui/pptx';

	import { useTranslator } from '../../../i18n/context';
	import type { EditorState } from '../../editor/editor-state.svelte';
	import InspectorFields from './InspectorFields.svelte';

	const { editor }: { editor: EditorState } = $props();
	const t = useTranslator();

	const slide = $derived(editor.slides[editor.currentSlideIndex]);
	const clearState = $derived(slideBackgroundClearState(slide, editor.editable));
</script>

<InspectorFields
	sections={SLIDE_BACKGROUND_SECTIONS}
	target={slide}
	disabled={!editor.editable}
	onpatch={(patch) => editor.backgroundOps.patchSlide(patch)}
	oncolorcommit={(hex) => editor.recordRecentColor(hex)}
/>
{#if clearState.visible}
	<button
		type="button"
		class="pptx-svelte-background-clear"
		disabled={!clearState.enabled}
		onclick={() => editor.backgroundOps.clearSlideBackground()}
	>
		{t('pptx.slideBackground.clearBackground')}
	</button>
{/if}

<style>
	.pptx-svelte-background-clear {
		width: 100%;
		min-height: 28px;
		margin-top: 8px;
		border: 1px solid var(--pptx-border, var(--office-border, #d1d5db));
		border-radius: 5px;
		background: var(--pptx-muted, var(--office-surface, #f3f4f6));
		color: inherit;
		font-size: 11px;
	}
</style>

<script lang="ts">
	/**
	 * ContextualTab: the body of a selection-driven tab (Shape Format, Picture
	 * Format, Table Design, Chart Design, SmartArt Design). Which groups it
	 * shows, and which galleries sit in them in which mode, is shared's
	 * `CONTEXTUAL_TAB_GROUPS`; this file only maps that list onto ribbon groups.
	 * A command group (Picture Format > Size) draws its shared strip as well.
	 */
	import { CONTEXTUAL_TAB_GROUPS } from 'ooxml-ui/pptx';
	import type { PictureRibbonState, RibbonContextualTabId, ToolbarActionId } from 'ooxml-ui/pptx';

	import { useTranslator } from '../../../../i18n/context';
	import type { EditorState } from '../../../editor/editor-state.svelte';
	import ChartSelectionStrip from './ChartSelectionStrip.svelte';
	import ArrangeExtras from '../home/ArrangeExtras.svelte';
	import ArrangeHomeStrip from '../home/ArrangeHomeStrip.svelte';
	import type { ChromeUiState } from '../../../state/chrome-ui.svelte';
	import PictureSizeStrip from './PictureSizeStrip.svelte';
	import RibbonGallery from './RibbonGallery.svelte';
	import { translatedOr } from './gallery-labels';

	const {
		tab,
		editor,
		hiddenActions,
		chromeUi,
		onselectionpane,
	}: {
		tab: RibbonContextualTabId;
		editor: EditorState;
		hiddenActions?: readonly ToolbarActionId[];
		chromeUi?: ChromeUiState | undefined;
		onselectionpane?: (() => void) | undefined;
	} = $props();
	const t = useTranslator();
	const groups = $derived(CONTEXTUAL_TAB_GROUPS[tab]);
	const picture = $derived<PictureRibbonState>({
		element: editor.selectedElement,
		editable: editor.editable,
		translate: t,
		slides: editor.slides,
		patch: (id, patch) => editor.applyElementPatch(id, patch),
		loadImage: (path) => editor.getHandler()?.getImageData(path) ?? Promise.resolve(undefined),
		updateSlides: (update) => editor.commitSlides(update(editor.slides)),
		altText: chromeUi && (() => chromeUi.formatPane.requestFormatPane('picture', { focus: true })),
		selectionPane: onselectionpane,
	});
</script>

<div class="pptx-svelte-ctxtab" data-ribbon-contextual-panel={tab}>
	{#each groups as group (group.group)}
		<pptx-ui-ribbon-group label={translatedOr(t, group.labelKey, group.label)} data-ribbon-group={group.group}>
			{#each group.galleries as placement (placement.control)}
				<RibbonGallery {placement} />
			{/each}
			{#if group.strip === 'picture-size'}
				<PictureSizeStrip {editor} {hiddenActions} />
			{:else if group.strip === 'chart-selection'}
				<ChartSelectionStrip {editor} />
			{/if}
			{#if group.pictureStrip === 'picture-arrange'}
				<ArrangeExtras {editor} {hiddenActions} />
				<ArrangeHomeStrip {editor} strip="order" />
			{/if}
			{#if group.pictureStrip}
				<pptx-ui-picture-ribbon family={group.pictureStrip} state={picture}></pptx-ui-picture-ribbon>
			{/if}
		</pptx-ui-ribbon-group>
	{/each}
</div>

<style>
	.pptx-svelte-ctxtab {
		display: flex;
		align-items: stretch;
		flex-wrap: nowrap;
	}
</style>

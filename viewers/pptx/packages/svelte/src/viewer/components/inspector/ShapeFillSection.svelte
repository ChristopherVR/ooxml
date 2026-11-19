<script lang="ts">
	/**
	 * ShapeFillSection: the Format pane's Fill section. The fill kind,
	 * transparency, gradient type and angle and picture tiling are the shared
	 * fill descriptors drawn by {@link InspectorFields}; the solid colour, the
	 * gradient stops ({@link GradientPanel}), the pattern swatches
	 * ({@link PatternFillPanel}) and the picture source are what
	 * `resolveShapeFillView` asks for, and their events are the shared
	 * `shapeFillPatch` actions. The solid colour picker previews every shade of
	 * a drag and commits it as one undo step (`live-color.ts`).
	 */
	import type { PptxElement, ShapeStyle } from 'pptx-viewer-core';
	import { hasShapeProperties } from 'pptx-viewer-core';
	import type { ShapeFillAction, ThemeColorPickerCommit } from 'ooxml-ui/pptx';
	import {
		fillColorOf,
		linkedContentUrl,
		resolveShapeFillView,
		SHAPE_FILL_MODE_SECTIONS,
		SHAPE_FILL_OPTIONS_SECTIONS,
		shapeFillPatch,
		shapeQuickStylePatch,
	} from 'ooxml-ui/pptx';

	import { useTranslator } from '../../../i18n/context';
	import type { EditorState } from '../../editor/editor-state.svelte';
	import { setSolidFillPatch } from '../../editor';
	import GradientPanel from './GradientPanel.svelte';
	import InspectorFields from './InspectorFields.svelte';
	import { getLinkedContentGetter } from '../../state/linked-content';
	import { liveColorEvents, liveElementColor } from './live-color';
	import PatternFillPanel from './PatternFillPanel.svelte';
	import RecentColorsRow from './RecentColorsRow.svelte';
	import ThemeColorSwatchGrid from './ThemeColorSwatchGrid.svelte';

	const { editor, el }: { editor: EditorState; el: PptxElement } = $props();
	const t = useTranslator();

	const style = $derived(hasShapeProperties(el) ? el.shapeStyle : undefined);
	const view = $derived(resolveShapeFillView(style));
	// The fill picture preview; the address field keeps showing the model's value.
	const getLinkedContent = getLinkedContentGetter();
	const imagePreview = $derived(linkedContentUrl(view.image?.url, getLinkedContent()));

	function patchStyle(patch: Partial<ShapeStyle>): void {
		editor.patchSelected(shapeQuickStylePatch(el, patch));
	}

	function act(action: ShapeFillAction): void {
		const patch = shapeFillPatch(style, action);
		if (patch) {
			patchStyle(patch);
		}
	}

	const remember = (hex: string): void => editor.recordRecentColor(hex);

	const fillInput = liveElementColor(
		() => editor,
		() => el,
		fillColorOf,
		(cur, hex) => setSolidFillPatch(cur, hex, undefined),
	);

	function commitFill(hex: string): void {
		act({ type: 'color', hex });
		remember(hex);
	}

	function commitFillTheme(commit: ThemeColorPickerCommit): void {
		act({ type: 'color', hex: commit.hex, ref: commit.ref });
		remember(commit.hex);
	}

	function onImageFile(input: HTMLInputElement): void {
		const file = input.files?.[0];
		if (!file) {
			return;
		}
		const reader = new FileReader();
		reader.onload = () => act({ type: 'imageUrl', url: String(reader.result) });
		reader.readAsDataURL(file);
	}
</script>

<InspectorFields sections={SHAPE_FILL_MODE_SECTIONS} target={style} onpatch={patchStyle} />

{#if view.color}
	<label class="pptx-svelte-inspector-color">
		<span>{t('pptx.inspector.fill')}</span>
		<input type="color" value={view.color.value} use:liveColorEvents={fillInput} />
		<ThemeColorSwatchGrid
			themeColorMap={editor.themeColorMap}
			selectedRef={view.color.colorRef}
			selectedHex={view.color.value}
			onpick={commitFillTheme}
		/>
		<RecentColorsRow colors={editor.mruColors} onselect={commitFill} />
	</label>
{/if}

<InspectorFields sections={SHAPE_FILL_OPTIONS_SECTIONS} target={style} onpatch={patchStyle} />

{#if view.stops}
	<GradientPanel
		stops={view.stops}
		themeColorMap={editor.themeColorMap}
		onaction={act}
		oncolorcommit={remember}
	/>
{/if}

{#if view.pattern}
	<PatternFillPanel pattern={view.pattern} onaction={act} oncolorcommit={remember} />
{/if}

{#if view.image}
	<div class="pptx-svelte-fill-image">
		<label>
			<span>{t('pptx.fillAdvanced.imageUrl')}</span>
			<input
				type="text"
				data-testid="fill-image-url"
				placeholder="https://example.com/image.png"
				value={view.image.url}
				onchange={(e) => act({ type: 'imageUrl', url: e.currentTarget.value })}
			/>
		</label>
		<label>
			<span>{t('pptx.fillAdvanced.orChooseFile')}</span>
			<input type="file" accept="image/*" onchange={(e) => onImageFile(e.currentTarget)} />
		</label>
		{#if imagePreview}
			<img src={imagePreview} alt={t('pptx.fillAdvanced.fillPreviewAlt')} />
		{/if}
	</div>
{/if}

<style>
	.pptx-svelte-inspector-color {
		display: flex;
		flex-direction: column;
		gap: 3px;
		margin-top: 8px;
	}

	.pptx-svelte-inspector-color span,
	.pptx-svelte-fill-image span {
		color: var(--pptx-muted-foreground, var(--office-muted-foreground, #6b7280));
	}

	.pptx-svelte-inspector-color input[type='color'] {
		width: 40px;
		height: 26px;
		padding: 0;
		border: 1px solid var(--pptx-border, var(--office-border, #d1d5db));
		border-radius: 4px;
		background: transparent;
		cursor: pointer;
	}

	.pptx-svelte-fill-image {
		display: flex;
		flex-direction: column;
		gap: 6px;
		margin-top: 10px;
	}

	.pptx-svelte-fill-image label {
		display: flex;
		flex-direction: column;
		gap: 3px;
	}

	.pptx-svelte-fill-image input[type='text'] {
		height: 26px;
		box-sizing: border-box;
		padding: 0 6px;
		border: 1px solid var(--pptx-border, var(--office-border, #d1d5db));
		border-radius: var(--pptx-radius, var(--office-radius, 4px));
		background: var(--pptx-background, var(--office-background, #ffffff));
		color: inherit;
		font: inherit;
	}

	.pptx-svelte-fill-image img {
		width: 100%;
		height: 64px;
		object-fit: cover;
		border: 1px solid var(--pptx-border, var(--office-border, #d1d5db));
		border-radius: 4px;
	}
</style>

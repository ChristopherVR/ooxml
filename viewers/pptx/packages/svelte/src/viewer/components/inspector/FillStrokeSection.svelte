<script lang="ts">
	/**
	 * FillStrokeSection: the shape's Fill ({@link ShapeFillSection}: the shared
	 * fill descriptors and actions) and its Line (the line colour with its
	 * theme grid and recent colours, then the shared Line options
	 * `SHAPE_LINE_SECTIONS` drawn by `InspectorFields`). Shown only for
	 * elements that pass `hasShapeProperties`. The line colour picker previews
	 * every shade of a drag and commits it as one undo step (`live-color.ts`).
	 */
	import type { PptxElement, ShapeStyle } from 'pptx-viewer-core';
	import { hasShapeProperties } from 'pptx-viewer-core';
	import type { ThemeColorPickerCommit } from 'ooxml-ui/pptx';
	import { SHAPE_LINE_SECTIONS, shapeQuickStylePatch, strokeColorOf } from 'ooxml-ui/pptx';

	import { useTranslator } from '../../../i18n/context';
	import type { EditorState } from '../../editor/editor-state.svelte';
	import { setStrokeColorPatch } from '../../editor';
	import InspectorFields from './InspectorFields.svelte';
	import { liveColorEvents, liveElementColor } from './live-color';
	import RecentColorsRow from './RecentColorsRow.svelte';
	import ShapeFillSection from './ShapeFillSection.svelte';
	import ThemeColorSwatchGrid from './ThemeColorSwatchGrid.svelte';

	const {
		editor,
		el,
		part,
	}: {
		editor: EditorState;
		el: PptxElement;
		/** The Format pane section drawn (`fill` or `outline`); omitted, both. */
		part?: 'fill' | 'outline';
	} = $props();
	const showFill = $derived(part !== 'outline');
	const showLine = $derived(part !== 'fill');
	const t = useTranslator();

	const stroke = $derived(strokeColorOf(el));
	const shapeStyle = $derived(hasShapeProperties(el) ? el.shapeStyle : undefined);

	const strokeInput = liveElementColor(
		() => editor,
		() => el,
		strokeColorOf,
		(cur, hex) => setStrokeColorPatch(cur, hex, undefined),
	);

	function commitStroke(hex: string): void {
		editor.patchSelected(setStrokeColorPatch(el, hex, undefined));
		editor.recordRecentColor(hex);
	}

	function commitStrokeTheme(commit: ThemeColorPickerCommit): void {
		editor.patchSelected(setStrokeColorPatch(el, commit.hex, commit.ref));
		editor.recordRecentColor(commit.hex);
	}

	const strokeColorRef = $derived(
		hasShapeProperties(el) ? el.shapeStyle?.strokeColorRef : undefined,
	);

	function patchLine(patch: Partial<ShapeStyle>): void {
		editor.patchSelected(shapeQuickStylePatch(el, patch));
	}
</script>

{#if showFill}
	<div class="pptx-svelte-inspector-fill" data-pptx-inspector-section="fill-stroke">
		<ShapeFillSection {editor} {el} />
	</div>
{/if}

{#if showLine}
	<label class="pptx-svelte-inspector-color">
		<span>{t('pptx.inspector.line')}</span>
		<input
			type="color"
			value={/^#/.test(stroke) ? stroke : '#000000'}
			use:liveColorEvents={strokeInput}
		/>
		<ThemeColorSwatchGrid
			themeColorMap={editor.themeColorMap}
			selectedRef={strokeColorRef}
			selectedHex={stroke}
			onpick={commitStrokeTheme}
		/>
		<RecentColorsRow colors={editor.mruColors} onselect={commitStroke} />
	</label>
	<InspectorFields sections={SHAPE_LINE_SECTIONS} target={shapeStyle} onpatch={patchLine} />
{/if}

<style>
	.pptx-svelte-inspector-fill {
		margin-bottom: 12px;
	}

	.pptx-svelte-inspector-color {
		display: flex;
		flex-direction: column;
		gap: 3px;
	}

	.pptx-svelte-inspector-color span {
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
</style>

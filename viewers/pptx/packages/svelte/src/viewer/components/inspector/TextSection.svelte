<script lang="ts">
	/**
	 * TextSection: the shared Text Box options (`TEXT_BOX_SECTIONS`: vertical
	 * anchor, direction, autofit mode, wrap-in-shape, columns, drawn by
	 * `InspectorFields`), text colour, text effects and paragraph spacing, for
	 * elements carrying text properties. Paragraph alignment (left/center/right/
	 * justify) is already covered by the Home tab's Paragraph group
	 * (`ParagraphGroup.svelte`), so it is intentionally not duplicated here.
	 * The text colour picker previews every shade of a drag and commits it as
	 * one undo step (`live-color.ts`).
	 */
	import type { PptxElement, PptxThemeColorRef, TextStyle } from 'pptx-viewer-core';
	import { TEXT_BOX_SECTIONS, textColorOf, textStylePatch } from 'ooxml-ui/pptx';
	import type { ThemeColorPickerCommit } from 'ooxml-ui/pptx';

	import { useTranslator } from '../../../i18n/context';
	import type { EditorState } from '../../editor/editor-state.svelte';
	import { setTextColorPatch } from '../../editor';
	import InspectorFields from './InspectorFields.svelte';
	import { liveColorEvents, liveElementColor } from './live-color';
	import RecentColorsRow from './RecentColorsRow.svelte';
	import TextEffectsSection from './TextEffectsSection.svelte';
	import ThemeColorSwatchGrid from './ThemeColorSwatchGrid.svelte';

	const { editor, el }: { editor: EditorState; el: PptxElement } = $props();
	const t = useTranslator();

	const textColor = $derived(textColorOf(el));
	const textColorRef = $derived('textStyle' in el ? el.textStyle?.colorRef : undefined);
	const textStyle = $derived('textStyle' in el ? (el.textStyle ?? {}) : {});

	/** Native colour input: always clears `colorRef` (theme-grid pick keeps it). */
	const textColorInput = liveElementColor(
		() => editor,
		() => el,
		textColorOf,
		(cur, hex) => setTextColorPatch(cur, hex, undefined),
	);

	/** A swatch or recent-colour pick: one committed edit. */
	function commitTextColor(hex: string, ref?: PptxThemeColorRef): void {
		editor.patchSelected(setTextColorPatch(el, hex, ref));
		editor.recordRecentColor(hex);
	}
	function commitTextColorTheme(commit: ThemeColorPickerCommit): void {
		commitTextColor(commit.hex, commit.ref);
	}
	function patchText(next: Partial<TextStyle>): void {
		editor.patchSelected(textStylePatch(el, next));
	}
</script>

<InspectorFields sections={TEXT_BOX_SECTIONS} target={textStyle} onpatch={patchText} />

<label class="pptx-svelte-field pptx-svelte-text-color-field">
	<span class="pptx-svelte-field-label">{t('pptx.textPanel.color')}</span>
	<input
		type="color"
		class="pptx-svelte-text-color"
		value={/^#/.test(textColor) ? textColor : '#000000'}
		use:liveColorEvents={textColorInput}
	/>
</label>
<ThemeColorSwatchGrid
	themeColorMap={editor.themeColorMap}
	selectedRef={textColorRef}
	selectedHex={textColor}
	onpick={commitTextColorTheme}
/>
<RecentColorsRow colors={editor.mruColors} onselect={commitTextColor} />

<TextEffectsSection {editor} {el} />

<div class="pptx-svelte-grid">
	<label class="pptx-svelte-field"><span>{t('pptx.textAdvanced.lineSpacing')}</span><input type="number" min="0.5" max="5" step="0.05" value={textStyle.lineSpacing ?? 1} onchange={(event) => patchText({ lineSpacing: Number(event.currentTarget.value), lineSpacingExactPt: undefined })} /></label>
	<label class="pptx-svelte-field"><span>{t('pptx.textAdvanced.spaceBefore')}</span><input type="number" min="0" value={textStyle.paragraphSpacingBefore ?? 0} onchange={(event) => patchText({ paragraphSpacingBefore: Number(event.currentTarget.value) })} /></label>
	<label class="pptx-svelte-field"><span>{t('pptx.textAdvanced.spaceAfter')}</span><input type="number" min="0" value={textStyle.paragraphSpacingAfter ?? 0} onchange={(event) => patchText({ paragraphSpacingAfter: Number(event.currentTarget.value) })} /></label>
</div>
<label class="pptx-svelte-field-checkbox"><pptx-ui-checkbox checked={textStyle.rtl ?? false} onchange={(event) => patchText({ rtl: event.currentTarget.checked })} ></pptx-ui-checkbox><span>{t('pptx.textAdvanced.rtl')}</span></label>

<style>
	.pptx-svelte-field {
		display: flex;
		flex-direction: column;
		gap: 3px;
		margin-top: 10px;
	}

	.pptx-svelte-field:first-child {
		margin-top: 0;
	}

	.pptx-svelte-field-label {
		color: var(--pptx-muted-foreground, var(--office-muted-foreground, #6b7280));
	}

	.pptx-svelte-field select {
		height: 26px;
		border: 1px solid var(--pptx-border, var(--office-border, #d1d5db));
		border-radius: var(--pptx-radius, var(--office-radius, 4px));
		background: var(--pptx-background, var(--office-background, #ffffff));
		color: inherit;
		font: inherit;
	}
	.pptx-svelte-field input { height: 26px; box-sizing: border-box; border: 1px solid var(--pptx-border, var(--office-border, #d1d5db)); border-radius: var(--pptx-radius, var(--office-radius, 4px)); background: var(--pptx-background, var(--office-background, #ffffff)); color: inherit; font: inherit; }
	.pptx-svelte-grid { display: grid; grid-template-columns: minmax(0, 1fr) minmax(0, 1fr); gap: 0 6px; }
	.pptx-svelte-grid input { min-width: 0; width: 100%; }

	.pptx-svelte-text-color-field {
		align-items: flex-start;
	}

	.pptx-svelte-text-color {
		width: 40px;
		padding: 0;
		cursor: pointer;
	}

	.pptx-svelte-field-checkbox {
		display: flex;
		align-items: center;
		gap: 6px;
		margin-top: 10px;
		cursor: pointer;
	}
</style>

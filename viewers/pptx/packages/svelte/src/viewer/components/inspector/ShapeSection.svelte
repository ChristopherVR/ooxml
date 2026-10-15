<script lang="ts">
	import type { PptxElement, ShapeStyle } from 'pptx-viewer-core';
	import { SHAPE_PRESET_DEFS } from 'ooxml-ui/pptx';

	import { useTranslator } from '../../../i18n/context';

	import type { EditorState } from '../../editor/editor-state.svelte';
	import ConnectorArrowsSection from './ConnectorArrowsSection.svelte';

	const { editor, el }: { editor: EditorState; el: PptxElement } = $props();
	const t = useTranslator();
	const style = $derived('shapeStyle' in el ? el.shapeStyle : undefined);
	const shapeType = $derived('shapeType' in el ? el.shapeType : undefined);

	function patchStyle(next: Partial<ShapeStyle>): void {
		editor.patchSelected({ shapeStyle: { ...style, ...next } } as Partial<PptxElement>);
	}
	function setShapeType(next: string): void {
		editor.patchSelected({ shapeType: next } as Partial<PptxElement>);
	}
	async function sampleFill(): Promise<void> {
		const Picker = (window as unknown as { EyeDropper?: new () => { open(): Promise<{ sRGBHex: string }> } }).EyeDropper;
		if (!Picker) {
			return;
		}
		const result = await new Picker().open();
		patchStyle({ fillMode: 'solid', fillColor: result.sRGBHex });
	}
</script>

{#if el.type === 'shape'}
	<label>{t('pptx.shape.type')}<pptx-ui-select aria-label={t('pptx.shape.type')} value={shapeType ?? 'rect'} onchange={(event) => setShapeType(event.currentTarget.value)}>{#each SHAPE_PRESET_DEFS as preset}<option value={preset.type}>{preset.label}</option>{/each}</pptx-ui-select></label>
{/if}
<button type="button" onclick={() => void sampleFill()} disabled={typeof window === 'undefined' || !('EyeDropper' in window)} title={t('pptx.fillStroke.eyedropperTooltip')}>{t('pptx.ribbon.eyedropper')}</button>
{#if el.type === 'connector'}
	<ConnectorArrowsSection {editor} {style} {el} />
{/if}

<style>label{display:grid;gap:3px;margin-top:7px;color:var(--pptx-muted-foreground);font-size:10px}select,button{height:26px;border:1px solid var(--pptx-border);border-radius:5px;background:var(--pptx-background);color:inherit}button{margin-top:8px}</style>

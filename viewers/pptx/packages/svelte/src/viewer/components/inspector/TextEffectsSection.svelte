<script lang="ts">
	import { applyTextStyleUpdate, buildTextWarpGallery, textWarpStyle } from 'ooxml-ui/pptx';
	import type { RibbonGalleryPickEvent } from 'ooxml-ui/pptx';
	import type { PptxElement, Text3DStyle, TextStyle } from 'pptx-viewer-core';
	import { hasTextProperties } from 'pptx-viewer-core';

	import { useTranslator } from '../../../i18n/context';
	import type { EditorState } from '../../editor/editor-state.svelte';
	import { strictTranslator } from '../ribbon/galleries/gallery-labels';
	import { bevelPresetLabel, materialPresetLabel } from './text-effects-labels';

	const { editor, el }: { editor: EditorState; el: PptxElement } = $props();
	const t = useTranslator();
	const style = $derived(hasTextProperties(el) ? (el.textStyle ?? {}) : {});
	// The element translates its tiles outside Svelte's tracking: read the title here so a language
	// switch hands it a new translator and the tiles are named again.
	const translateLabel = $derived.by(() => {
		void t('pptx.textWarp.title');
		return strictTranslator(t);
	});
	/** The shared Text Warp gallery: every OOXML preset, the element's own one marked. */
	const warpGallery = $derived(buildTextWarpGallery(style.textWarpPreset));
	/**
	 * The material / bevel values this panel offers. The shared preset
	 * catalogues list more of each; keeping these lists here means labelling
	 * them cannot change what the selects carry.
	 */
	const materials: readonly string[] = ['matte', 'plastic', 'metal', 'warmMatte', 'softEdge', 'flat'];
	const bevels: readonly string[] = ['none', 'circle', 'relaxedInset', 'coolSlant', 'angle', 'softRound', 'convex'];

	function patch(next: Partial<TextStyle>): void {
		editor.patchSelected({ textStyle: { ...style, ...next } } as Partial<PptxElement>);
	}
	/**
	 * Shadow, glow and reflection are run properties (`a:rPr/a:effectLst`), and a run is painted
	 * from its own style: they go to the body AND every run (the ribbon's whole-text update, as
	 * React's inspector applies them), or the slide would not show them.
	 */
	function patchRuns(next: Partial<TextStyle>): void {
		editor.patchSelected(applyTextStyleUpdate(el, next, null).patch);
	}

	/** A Text Warp pick sets its preset on the text style; No Transform clears it. */
	function pickWarp(event: Event): void {
		const next = textWarpStyle((event as RibbonGalleryPickEvent).detail.itemId);
		if (next) {
			patch(next);
		}
	}

	function patch3d(next: Partial<Text3DStyle>): void {
		patch({ text3d: { ...style.text3d, ...next } });
	}
	function pushColor(hex: string): void {
		editor.recordRecentColor(hex);
	}
</script>

<details>
	<summary>{t('pptx.textEffects.title')}</summary>
	<label class="check"><pptx-ui-checkbox checked={Boolean(style.textShadowColor)} onchange={(event) => patchRuns(event.currentTarget.checked ? { textShadowColor: '#000000', textShadowBlur: 4, textShadowOffsetX: 2, textShadowOffsetY: 2, textShadowOpacity: 0.5 } : { textShadowColor: undefined, textShadowBlur: undefined, textShadowOffsetX: undefined, textShadowOffsetY: undefined, textShadowOpacity: undefined })} ></pptx-ui-checkbox>{t('pptx.textEffects.shadow')}</label>
	{#if style.textShadowColor}<div class="grid"><label>{t('pptx.textEffects.color')}<input type="color" value={style.textShadowColor} onchange={(event) => { patchRuns({ textShadowColor: event.currentTarget.value }); pushColor(event.currentTarget.value); }} /></label><label>{t('pptx.textEffects.opacity')}<input type="number" min="0" max="1" step="0.05" value={style.textShadowOpacity ?? 0.5} onchange={(event) => patchRuns({ textShadowOpacity: Number(event.currentTarget.value) })} /></label><label>{t('pptx.textEffects.blur')}<input type="number" min="0" max="50" value={style.textShadowBlur ?? 4} onchange={(event) => patchRuns({ textShadowBlur: Number(event.currentTarget.value) })} /></label><label>{t('pptx.textEffects.offsetX')}<input type="number" min="-50" max="50" value={style.textShadowOffsetX ?? 2} onchange={(event) => patchRuns({ textShadowOffsetX: Number(event.currentTarget.value) })} /></label><label>{t('pptx.textEffects.offsetY')}<input type="number" min="-50" max="50" value={style.textShadowOffsetY ?? 2} onchange={(event) => patchRuns({ textShadowOffsetY: Number(event.currentTarget.value) })} /></label></div>{/if}
	<label class="check"><pptx-ui-checkbox checked={Boolean(style.textGlowColor)} onchange={(event) => patchRuns(event.currentTarget.checked ? { textGlowColor: '#ffff00', textGlowRadius: 6, textGlowOpacity: 0.6 } : { textGlowColor: undefined, textGlowRadius: undefined, textGlowOpacity: undefined })} ></pptx-ui-checkbox>{t('pptx.textEffects.glow')}</label>
	{#if style.textGlowColor}<div class="grid"><label>{t('pptx.textEffects.color')}<input type="color" value={style.textGlowColor} onchange={(event) => { patchRuns({ textGlowColor: event.currentTarget.value }); pushColor(event.currentTarget.value); }} /></label><label>{t('pptx.textEffects.opacity')}<input type="number" min="0" max="1" step="0.05" value={style.textGlowOpacity ?? 0.6} onchange={(event) => patchRuns({ textGlowOpacity: Number(event.currentTarget.value) })} /></label><label>{t('pptx.effects.radius')}<input type="number" min="1" max="50" value={style.textGlowRadius ?? 6} onchange={(event) => patchRuns({ textGlowRadius: Number(event.currentTarget.value) })} /></label></div>{/if}
	<label class="check"><pptx-ui-checkbox checked={style.textReflection ?? false} onchange={(event) => patchRuns(event.currentTarget.checked ? { textReflection: true, textReflectionBlur: 1, textReflectionOffset: 3, textReflectionStartOpacity: 0.5, textReflectionEndOpacity: 0 } : { textReflection: undefined, textReflectionBlur: undefined, textReflectionOffset: undefined, textReflectionStartOpacity: undefined, textReflectionEndOpacity: undefined })} ></pptx-ui-checkbox>{t('pptx.textEffects.reflection')}</label>
	{#if style.textReflection}<div class="grid"><label>{t('pptx.textEffects.blur')}<input type="number" min="0" max="20" value={style.textReflectionBlur ?? 1} onchange={(event) => patchRuns({ textReflectionBlur: Number(event.currentTarget.value) })} /></label><label>{t('pptx.textEffects.offset')}<input type="number" min="0" max="20" value={style.textReflectionOffset ?? 3} onchange={(event) => patchRuns({ textReflectionOffset: Number(event.currentTarget.value) })} /></label><label>{t('pptx.textEffects.startOpacity')}<input type="number" min="0" max="1" step="0.05" value={style.textReflectionStartOpacity ?? 0.5} onchange={(event) => patchRuns({ textReflectionStartOpacity: Number(event.currentTarget.value) })} /></label><label>{t('pptx.textEffects.endOpacity')}<input type="number" min="0" max="1" step="0.05" value={style.textReflectionEndOpacity ?? 0} onchange={(event) => patchRuns({ textReflectionEndOpacity: Number(event.currentTarget.value) })} /></label></div>{/if}
	<!-- Not a <label>: a click on its caption would press the first tile (No Transform). -->
	<div class="warp"><span>{t('pptx.textEffects.transform')}</span><pptx-ui-ribbon-gallery mode="panel" compact descriptor={warpGallery} {translateLabel} disabled={!editor.editable} ongallery-pick={pickWarp}></pptx-ui-ribbon-gallery></div>
	{#if style.textWarpPreset}<div class="grid"><label>{t('pptx.textEffects.adjustment')}<input type="number" value={style.textWarpAdj ?? 50000} onchange={(event) => patch({ textWarpAdj: Number(event.currentTarget.value) })} /></label><label>{t('pptx.textEffects.adjustment2')}<input type="number" value={style.textWarpAdj2 ?? 50000} onchange={(event) => patch({ textWarpAdj2: Number(event.currentTarget.value) })} /></label></div>{/if}
	<label class="check"><pptx-ui-checkbox checked={Boolean(style.text3d)} onchange={(event) => patch({ text3d: event.currentTarget.checked ? { extrusionHeight: 57150, extrusionColor: '#808080', presetMaterial: 'plastic', bevelTopType: 'circle', bevelTopWidth: 25400, bevelTopHeight: 25400 } : undefined })} ></pptx-ui-checkbox>{t('pptx.text3d.title')}</label>
	{#if style.text3d}<div class="grid"><label>{t('pptx.text3d.depth')}<input type="number" min="0" value={style.text3d.extrusionHeight ?? 0} onchange={(event) => patch3d({ extrusionHeight: Number(event.currentTarget.value) })} /></label><label>{t('pptx.text3d.extrusionColor')}<input type="color" value={style.text3d.extrusionColor ?? '#808080'} onchange={(event) => { patch3d({ extrusionColor: event.currentTarget.value }); pushColor(event.currentTarget.value); }} /></label><label>{t('pptx.text3d.material')}<pptx-ui-select aria-label={t('pptx.text3d.material')} value={style.text3d.presetMaterial ?? 'plastic'} onchange={(event) => patch3d({ presetMaterial: event.currentTarget.value as Text3DStyle['presetMaterial'] })}>{#each materials as material}<option value={material}>{materialPresetLabel(material, t)}</option>{/each}</pptx-ui-select></label><label>{t('pptx.text3d.bevelType')}<pptx-ui-select aria-label={t('pptx.text3d.bevelType')} value={style.text3d.bevelTopType ?? 'circle'} onchange={(event) => patch3d({ bevelTopType: event.currentTarget.value as Text3DStyle['bevelTopType'] })}>{#each bevels as bevel}<option value={bevel}>{bevelPresetLabel(bevel, t)}</option>{/each}</pptx-ui-select></label></div>{/if}
</details>

<style>details{margin-top:10px;border-top:1px solid var(--pptx-border, var(--office-border, #d1d5db));padding-top:8px}summary{cursor:pointer;font-weight:600}.check{display:flex;align-items:center;gap:6px;margin-top:8px}.grid{display:grid;grid-template-columns:1fr 1fr;gap:6px;margin:6px 0 0 16px}label,.warp{display:grid;gap:3px;margin-top:7px}label,.warp>span{color:var(--pptx-muted-foreground, var(--office-muted-foreground, #6b7280));font-size:10px}input,select{min-width:0;height:26px;border:1px solid var(--pptx-border, var(--office-border, #d1d5db));border-radius:5px;background:var(--pptx-background, var(--office-background, #ffffff));color:inherit}</style>

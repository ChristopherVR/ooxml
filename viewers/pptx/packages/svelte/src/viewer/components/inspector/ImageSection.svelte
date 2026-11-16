<script lang="ts">
	/**
	 * ImageSection: the picture card of an `isImageLikeElement` element. Its
	 * adjustments and crop insets are the shared descriptors
	 * (`PICTURE_ADJUSTMENT_SECTIONS`, `PICTURE_CROP_SECTIONS`) drawn by the
	 * generic `InspectorFields`; the compound picture controls stay native
	 * markup here, and `AltTextSection` draws the alt text after it.
	 */
	import type { PptxElement } from 'pptx-viewer-core';
	import {
		cropResetPatch,
		cropResetState,
		imageResetPatch,
		imageResetState,
		ARTISTIC_EFFECTS,
		DUOTONE_PRESETS,
		PICTURE_ADJUSTMENT_SECTIONS,
		PICTURE_CROP_SECTIONS,
	} from 'ooxml-ui/pptx';

	import { useTranslator } from '../../../i18n/context';
	import type { EditorState } from '../../editor/editor-state.svelte';
	import InspectorFields from './InspectorFields.svelte';

	const { editor, el }: { editor: EditorState; el: PptxElement } = $props();
	const t = useTranslator();

	// G7: `a:picLocks/@noCrop` forbids cropping this specific picture.
	const croppable = $derived(cropResetState(el, editor.editable, true).enabled);
	const resetPictureEnabled = $derived(imageResetState(el, editor.editable, true).enabled);
	const effects = $derived('imageEffects' in el ? el.imageEffects : undefined);
	function setEffects(next: Record<string, unknown>): void {
		editor.patchSelected({ imageEffects: { ...effects, ...next } } as Partial<PptxElement>);
	}

	function replaceImage(file: File | undefined): void {
		if (!file) {
			return;
		}
		const reader = new FileReader();
		reader.onload = () => editor.patchSelected({ imageData: String(reader.result), imagePath: undefined } as Partial<PptxElement>);
		reader.readAsDataURL(file);
	}
	function resetImage(): void {
		if (resetPictureEnabled) {
			editor.patchSelected(imageResetPatch());
		}
	}
	function resetCrop(): void {
		if (croppable) {
			editor.patchSelected(cropResetPatch());
		}
	}
</script>

<div class="pptx-svelte-image-actions"><label>{t('pptx.image.replaceImage')}<input type="file" accept="image/*" onchange={(event) => replaceImage(event.currentTarget.files?.[0])} /></label><button type="button" disabled={!resetPictureEnabled} onclick={resetImage}>{t('pptx.image.resetImage')}</button></div>

<InspectorFields
	sections={PICTURE_ADJUSTMENT_SECTIONS}
	target={el}
	disabled={!editor.editable}
	onpatch={(patch) => editor.patchSelected(patch)}
/>
<InspectorFields
	sections={PICTURE_CROP_SECTIONS}
	target={el}
	disabled={!croppable}
	onpatch={(patch) => {
		if (croppable) {
			editor.patchSelected(patch);
		}
	}}
/>

<button type="button" class="pptx-svelte-reset-crop" disabled={!croppable} onclick={resetCrop}>{t('pptx.image.resetCrop')}</button>

<label class="pptx-svelte-field"><span>{t('pptx.image.artisticEffects')}</span><pptx-ui-select aria-label={t('pptx.image.artisticEffects')} value={effects?.artisticEffect ?? 'none'} onchange={(e) => setEffects({ artisticEffect: e.currentTarget.value === 'none' ? undefined : e.currentTarget.value })}>{#each ARTISTIC_EFFECTS as preset}<option value={preset[0]}>{t(preset[1])}</option>{/each}</pptx-ui-select></label>
<div class="pptx-svelte-duotone"><span>{t('pptx.image.duotone')}</span>{#each DUOTONE_PRESETS as preset}<button type="button" title={t(preset.labelKey)} style={`--shadow:${preset.shadow};--highlight:${preset.highlight}`} onclick={() => setEffects({ duotone: { color1: preset.shadow, color2: preset.highlight } })}></button>{/each}<button type="button" title={t('pptx.image.duotoneClear')} onclick={() => setEffects({ duotone: undefined })}>×</button></div>
<label class="pptx-svelte-field-checkbox"><pptx-ui-checkbox checked={Boolean(effects?.colorWash)} onchange={(event) => setEffects({ colorWash: event.currentTarget.checked ? { color: '#0066cc', opacity: 40 } : undefined })} ></pptx-ui-checkbox><span>{t('pptx.image.colorWash')}</span></label>
{#if effects?.colorWash}<div class="pptx-svelte-inspector-grid"><label><span>{t('pptx.image.washColor')}</span><input type="color" value={effects.colorWash.color} onchange={(event) => { setEffects({ colorWash: { ...effects.colorWash, color: event.currentTarget.value } }); editor.recordRecentColor(event.currentTarget.value); }} /></label><label><span>{t('pptx.image.washOpacity')}</span><input type="number" min="0" max="100" value={effects.colorWash.opacity ?? 40} onchange={(event) => setEffects({ colorWash: { ...effects.colorWash, opacity: Number(event.currentTarget.value) } })} /></label></div>{/if}

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

	.pptx-svelte-field-label b {
		color: inherit;
		font-weight: 600;
	}

	.pptx-svelte-field input[type='range'] {
		width: 100%;
	}

	.pptx-svelte-inspector-grid {
		display: grid;
		grid-template-columns: 1fr 1fr;
		gap: 8px;
		margin-top: 10px;
	}

	.pptx-svelte-inspector-grid label {
		display: flex;
		flex-direction: column;
		gap: 3px;
	}

	.pptx-svelte-inspector-grid span {
		color: var(--pptx-muted-foreground, var(--office-muted-foreground, #6b7280));
	}

	.pptx-svelte-inspector-grid input {
		width: 100%;
		height: 26px;
		box-sizing: border-box;
		padding: 0 6px;
		border: 1px solid var(--pptx-border, var(--office-border, #d1d5db));
		border-radius: var(--pptx-radius, var(--office-radius, 4px));
		background: var(--pptx-background, var(--office-background, #ffffff));
		color: inherit;
		font: inherit;
	}
	.pptx-svelte-field select{height:26px;border:1px solid var(--pptx-border, var(--office-border, #d1d5db));border-radius:6px;background:var(--pptx-background, var(--office-background, #ffffff));color:inherit}.pptx-svelte-duotone{display:flex;align-items:center;flex-wrap:wrap;gap:4px;margin-top:10px}.pptx-svelte-duotone>span{width:100%;color:var(--pptx-muted-foreground, var(--office-muted-foreground, #6b7280));font-size:10px}.pptx-svelte-duotone button{width:24px;height:24px;border:1px solid var(--pptx-border, var(--office-border, #d1d5db));border-radius:50%;background:linear-gradient(135deg,var(--shadow) 50%,var(--highlight) 50%);color:inherit}
	.pptx-svelte-image-actions{display:grid;grid-template-columns:1fr 1fr;gap:6px}.pptx-svelte-image-actions label,.pptx-svelte-image-actions button{display:grid;place-items:center;min-height:28px;border:1px solid var(--pptx-border, var(--office-border, #d1d5db));border-radius:5px;background:var(--pptx-muted, var(--office-surface, #f3f4f6));color:inherit;font-size:10px}.pptx-svelte-image-actions input{position:absolute;width:1px;height:1px;opacity:0}.pptx-svelte-reset-crop{width:100%;min-height:28px;margin-top:8px;border:1px solid var(--pptx-border, var(--office-border, #d1d5db));border-radius:5px;background:var(--pptx-muted, var(--office-surface, #f3f4f6));color:inherit;font-size:10px}.pptx-svelte-field-checkbox{display:flex;align-items:center;gap:5px;margin-top:8px}
</style>

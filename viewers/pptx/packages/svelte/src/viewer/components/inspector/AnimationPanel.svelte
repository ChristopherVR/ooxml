<script lang="ts">
	/**
	 * AnimationPanel: the docked per-element animation editor, port of React's
	 * `inspector/AnimationPanel.tsx`. Docked at the bottom of the inspector
	 * pane whenever an element is selected (any tab), it edits the slide-level
	 * `PptxSlide.animations` entry for the selection: entrance / emphasis /
	 * exit presets, direction (for directional presets), sequence, the TIMING
	 * block ({@link AnimationTimingFields}), and the timeline section
	 * ({@link AnimationTimelineSection}).
	 *
	 * Every mutation goes through the shared `animation-authoring` setters and
	 * commits via `commitSlideAnimations` -> `EditorState.commitSlides`, so each
	 * change is one undoable history step, exactly like React's
	 * `onUpdateSlide({ animations })` path. Self-gating: renders nothing
	 * without a selection, so the host only adds `<AnimationPanel {editor} />`.
	 */
	import Play from '@lucide/svelte/icons/play';
	import type {
		PptxAfterAnimationAction,
		PptxAnimationDirection,
		PptxAnimationPreset,
		PptxAnimationSequence,
		PptxElementAnimation,
	} from 'pptx-viewer-core';
	import {
	getAfterAnimationState,
		animationFor,
		applyMotionPathPreset,
		clearMotionPath,
		getDensePanelTouchTargetPx,
		getEffectSoundState,
		hasAnimation,
		setAfterAnimation,
		setAfterAnimationColor,
		setAnimationEmphasis,
		setAnimationEntrance,
		setAnimationExit,
		directionValuesFor,
		effectiveDirection,
		setDirection,
		setEffectSound,
		setEffectStockSound,
		setSequence,
		showDirectionPicker,
	} from 'ooxml-ui/pptx';

	import { useTranslator } from '../../../i18n/context';
	import type { EditorState } from '../../editor/editor-state.svelte';
	import { useWindowViewport } from '../../state/window-viewport.svelte';
	import AfterAnimationRow from './AfterAnimationRow.svelte';
	import { commitSlideAnimations } from './animation-panel-helpers';
	import {
		PANEL_DIRECTION_OPTIONS,
		PANEL_EMPHASIS_PRESETS,
		PANEL_ENTRANCE_PRESETS,
		PANEL_EXIT_PRESETS,
		PANEL_SEQUENCE_OPTIONS,
	} from './animation-panel-options';
	import { startAnimationPreview } from './animation-preview-control';
	import AnimationTimelineSection from './AnimationTimelineSection.svelte';
	import AnimationTimingFields from './AnimationTimingFields.svelte';
	import EffectSoundRow from './EffectSoundRow.svelte';
	import MotionPathRow from './MotionPathRow.svelte';

	const { editor }: { editor: EditorState } = $props();
	const t = useTranslator();
	const viewport = useWindowViewport();
	// The direction-picker icon buttons and the Preview button are 28px by
	// design; below the dense-panel breakpoint they must still clear the 44px
	// touch target, so this only ever grows them, never shrinks below 28.
	const touchPx = $derived(Math.max(28, getDensePanelTouchTargetPx(viewport.width)));
	const touchStyle = $derived(`min-width: ${touchPx}px; min-height: ${touchPx}px;`);

	const slide = $derived(editor.slides[editor.currentSlideIndex]);
	const el = $derived(editor.selectedElement);
	const anims = $derived<PptxElementAnimation[]>(slide?.animations ?? []);
	const anim = $derived(el ? animationFor(anims, el.id) : undefined);
	const hasAnim = $derived(el ? hasAnimation(anims, el.id) : false);
	const showDirection = $derived(el ? showDirectionPicker(anims, el.id) : false);
	// Only the directions PowerPoint has a variant for (shared catalogue).
	const directionValues = $derived(el ? directionValuesFor(anims, el.id) : []);
	const directionOptions = $derived(PANEL_DIRECTION_OPTIONS.filter((option) => directionValues.includes(option.value)));
	const activeDirection = $derived(effectiveDirection(anim, directionValues));
	const canEdit = $derived(editor.editable);
	const soundState = $derived(el ? getEffectSoundState(anims, el.id) : { hasSound: false });

	function commit(next: PptxElementAnimation[]): void {
		if (canEdit) {
			commitSlideAnimations(editor, next);
		}
	}

	function onEffect(group: 'entrance' | 'emphasis' | 'exit', value: string): void {
		if (!el) {
			return;
		}
		const preset = value as PptxAnimationPreset | 'none';
		const apply =
			group === 'entrance'
				? setAnimationEntrance
				: group === 'emphasis'
					? setAnimationEmphasis
					: setAnimationExit;
		commit(apply(anims, el.id, preset));
	}

	function onPreview(): void {
		if (anim) {
			startAnimationPreview(anim);
		}
	}

	/**
	 * Motion path is geometry, not a preset, so it never routes through
	 * `onEffect`. `custom` is the read-only marker for a hand-dragged path:
	 * re-selecting it must not snap the path back to a catalogue entry.
	 */
	function onMotionPath(presetId: string): void {
		if (!el || presetId === 'custom') {
			return;
		}
		commit(
			presetId === 'none'
				? clearMotionPath(anims, el.id)
				: applyMotionPathPreset(anims, el.id, presetId),
		);
	}

	function onEffectSoundPick(pick: { dataUrl: string; fileName?: string } | undefined): void {
		if (!el) {
			return;
		}
		commit(setEffectSound(anims, el.id, pick));
	}

	function onEffectStockSoundPick(catalogueId: string): void {
		if (!el) {
			return;
		}
		commit(setEffectStockSound(anims, el.id, catalogueId));
	}

	function onAfterAnimationChange(action: PptxAfterAnimationAction): void {
		if (!el) {
			return;
		}
		commit(setAfterAnimation(anims, el.id, action));
	}

	function onAfterAnimationColorChange(color: string): void {
		editor.recordRecentColor(color);
		if (!el) {
			return;
		}
		commit(setAfterAnimationColor(anims, el.id, color));
	}
</script>

{#if el && slide}
	<div class="pptx-svelte-animp" data-pptx-animation-panel>
		<div class="pptx-svelte-animp-header">
			<span class="pptx-svelte-animp-title">{t('pptx.animation.title')}</span>
			{#if hasAnim}
				<button
					type="button"
					class="pptx-svelte-animp-preview"
					style={touchStyle}
					title={t('pptx.animation.preview')}
					onclick={onPreview}
				>
					<Play size={12} aria-hidden="true" /> {t('pptx.animation.preview')}
				</button>
			{/if}
		</div>

		<label>
			<span>{t('pptx.animation.entrance')}</span>
			<pptx-ui-select aria-label={t('pptx.animation.entrance')} class="pptx-svelte-animp-entrance" disabled={!canEdit} value={anim?.entrance ?? 'none'} onchange={(e) => onEffect('entrance', e.currentTarget.value)}>
				<option value="none">{t('pptx.animation.none')}</option>
				{#each PANEL_ENTRANCE_PRESETS as preset (preset)}<option value={preset}>{t(`pptx.animation.preset.${preset}`)}</option>{/each}
			</pptx-ui-select>
		</label>

		<label>
			<span>{t('pptx.animation.emphasis')}</span>
			<pptx-ui-select aria-label={t('pptx.animation.emphasis')} class="pptx-svelte-animp-emphasis" disabled={!canEdit} value={anim?.emphasis ?? 'none'} onchange={(e) => onEffect('emphasis', e.currentTarget.value)}>
				<option value="none">{t('pptx.animation.none')}</option>
				{#each PANEL_EMPHASIS_PRESETS as preset (preset)}<option value={preset}>{t(`pptx.animation.preset.${preset}`)}</option>{/each}
			</pptx-ui-select>
		</label>

		<label>
			<span>{t('pptx.animation.exit')}</span>
			<pptx-ui-select aria-label={t('pptx.animation.exit')} class="pptx-svelte-animp-exit" disabled={!canEdit} value={anim?.exit ?? 'none'} onchange={(e) => onEffect('exit', e.currentTarget.value)}>
				<option value="none">{t('pptx.animation.none')}</option>
				{#each PANEL_EXIT_PRESETS as preset (preset)}<option value={preset}>{t(`pptx.animation.preset.${preset}`)}</option>{/each}
			</pptx-ui-select>
		</label>

		<!-- Motion path: geometry, not a preset, so it gets its own row. -->
		<MotionPathRow motionPath={anim?.motionPath} {canEdit} onchange={onMotionPath} />

		{#if hasAnim}
			{#if showDirection}
				<div class="pptx-svelte-animp-direction">
					<span>{t('pptx.animation.direction')}</span>
					<div class="pptx-svelte-animp-direction-row">
						{#each directionOptions as option (option.value)}
							<button
								type="button"
								disabled={!canEdit}
								style={touchStyle}
								class:is-active={activeDirection === option.value}
								title={t(option.labelKey)}
								aria-label={t(option.labelKey)}
								onclick={() => el && commit(setDirection(anims, el.id, option.value as PptxAnimationDirection))}
							>{option.glyph}</button>
						{/each}
					</div>
				</div>
			{/if}

			<label>
				<span>{t('pptx.animation.sequence')}</span>
				<pptx-ui-select aria-label={t('pptx.animation.sequence')} class="pptx-svelte-animp-sequence" disabled={!canEdit} value={anim?.sequence ?? 'asOne'} onchange={(e) => el && commit(setSequence(anims, el.id, e.currentTarget.value as PptxAnimationSequence))}>
					{#each PANEL_SEQUENCE_OPTIONS as option (option.value)}<option value={option.value}>{t(option.labelKey)}</option>{/each}
				</pptx-ui-select>
			</label>

			<EffectSoundRow
			{soundState}
			{canEdit}
			onpick={onEffectSoundPick}
			onpickstock={onEffectStockSoundPick}
		/>


			<AnimationTimingFields {editor} />
		{/if}

		{#if hasAnim || anim?.nativeAfterAnimation}
<AfterAnimationRow
				action={getAfterAnimationState(anim).action}
				color={getAfterAnimationState(anim).color}
			showColor={getAfterAnimationState(anim).showColor}
				{canEdit}
				onaction={onAfterAnimationChange}
				oncolor={onAfterAnimationColorChange}
			/>
{/if}
<AnimationTimelineSection {editor} selectedElementId={el.id} />
	</div>
{/if}

<style>
	.pptx-svelte-animp {
		flex: none;
		max-height: 220px;
		overflow-y: auto;
		padding: 8px 12px 12px;
		border-top: 1px solid var(--pptx-border, var(--office-border, #d1d5db));
		display: flex;
		flex-direction: column;
		gap: 8px;
		font-size: 11px;
	}

	.pptx-svelte-animp-header {
		display: flex;
		align-items: center;
		justify-content: space-between;
	}

	.pptx-svelte-animp-title {
		font-size: 11px;
		font-weight: 600;
		text-transform: uppercase;
		letter-spacing: 0.04em;
		color: var(--pptx-muted-foreground, var(--office-muted-foreground, #6b7280));
	}

	.pptx-svelte-animp-preview {
		display: inline-flex;
		align-items: center;
		gap: 4px;
		border: none;
		background: transparent;
		color: var(--pptx-primary, var(--office-accent, #2563eb));
		cursor: pointer;
		font: inherit;
		font-size: 10px;
	}

	.pptx-svelte-animp label {
		display: flex;
		flex-direction: column;
		gap: 3px;
	}

	.pptx-svelte-animp label > span,
	.pptx-svelte-animp-direction > span {
		color: var(--pptx-muted-foreground, var(--office-muted-foreground, #6b7280));
	}

	.pptx-svelte-animp select {
		width: 100%;
		height: 26px;
		box-sizing: border-box;
		padding: 0 6px;
		border: 1px solid var(--pptx-border, var(--office-border, #d1d5db));
		border-radius: var(--pptx-radius, var(--office-radius, 4px));
		background: var(--pptx-muted, var(--office-surface, #f3f4f6));
		color: inherit;
		font: inherit;
		font-size: 11px;
	}

	.pptx-svelte-animp-direction {
		display: flex;
		flex-direction: column;
		gap: 4px;
		padding-top: 6px;
		border-top: 1px solid var(--pptx-border, var(--office-border, #d1d5db));
	}

	.pptx-svelte-animp-direction-row {
		display: flex;
		gap: 4px;
	}

	.pptx-svelte-animp-direction-row button {
		display: grid;
		place-items: center;
		width: 28px;
		height: 28px;
		border: 1px solid var(--pptx-border, var(--office-border, #d1d5db));
		border-radius: var(--pptx-radius, var(--office-radius, 4px));
		background: var(--pptx-muted, var(--office-surface, #f3f4f6));
		color: var(--pptx-muted-foreground, var(--office-muted-foreground, #6b7280));
		cursor: pointer;
		font: inherit;
	}

	.pptx-svelte-animp-direction-row button.is-active {
		border-color: var(--pptx-primary, var(--office-accent, #2563eb));
		background: color-mix(in srgb, var(--pptx-primary, var(--office-accent, #2563eb)) 20%, transparent);
		color: var(--pptx-primary, var(--office-accent, #2563eb));
	}

	.pptx-svelte-animp-direction-row button:disabled {
		opacity: 0.5;
		cursor: default;
	}
</style>

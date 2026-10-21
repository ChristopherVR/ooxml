<script lang="ts">
	/**
	 * AfterAnimationRow: the animation panel's "after animation" row (dim to
	 * colour / hide after animation / hide on next click / don't dim), Svelte
	 * port of React's `inspector/AfterAnimationRow.tsx`.
	 */
	import type { PptxAfterAnimationAction } from 'pptx-viewer-core';
	import { AFTER_ANIMATION_VALUES } from 'ooxml-ui/pptx';

	import { useTranslator } from '../../../i18n/context';

	const {
		action,
		color,
		canEdit,
		showColor = true,
		onaction,
		oncolor,
	}: {
		action: PptxAfterAnimationAction;
		color: string | undefined;
		canEdit: boolean;
		showColor?: boolean;
		onaction: (action: PptxAfterAnimationAction) => void;
		oncolor: (color: string) => void;
	} = $props();

	const t = useTranslator();
</script>

<div class="pptx-svelte-after-animation">
	<label>
		<span>{t('pptx.animation.afterAnimation')}</span>
		<pptx-ui-select
			aria-label={t('pptx.animation.afterAnimation')}
			class="pptx-svelte-animp-after-animation"
			disabled={!canEdit}
			value={action}
			onchange={(event) => onaction(event.currentTarget.value as PptxAfterAnimationAction)}
		>
			{#each AFTER_ANIMATION_VALUES as value (value)}
				<option {value}>{t(`pptx.animation.afterAnimation.${value}`)}</option>
			{/each}
		</pptx-ui-select>
	</label>
	{#if action === 'dimToColor' && showColor}
		<label class="pptx-svelte-after-animation-color">
			<span>{t('pptx.animation.afterAnimation.color')}</span>
			<input
				type="color"
				aria-label={t('pptx.animation.afterAnimation.color')}
				disabled={!canEdit}
				value={color ?? '#808080'}
				onchange={(event) => oncolor(event.currentTarget.value)}
			/>
		</label>
	{/if}
</div>

<style>
	.pptx-svelte-after-animation {
		display: flex;
		flex-direction: column;
		gap: 6px;
	}

	.pptx-svelte-after-animation label {
		display: flex;
		flex-direction: column;
		gap: 3px;
	}

	.pptx-svelte-after-animation > label > span {
		color: var(--pptx-muted-foreground, var(--office-muted-foreground, #6b7280));
	}

	.pptx-svelte-animp-after-animation {
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

	.pptx-svelte-after-animation-color {
		flex-direction: row;
		align-items: center;
		gap: 6px;
	}

	.pptx-svelte-after-animation-color input[type='color'] {
		width: 40px;
		height: 24px;
		padding: 0;
	}
</style>

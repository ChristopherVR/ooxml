<script lang="ts">
	/**
	 * PatternFillPanel: the `a:pattFill` part of the Fill section, drawn from
	 * the shared pattern view (`resolveShapeFillView`): the 56-preset swatch
	 * grid in the shape's own colours, then the foreground and background
	 * colour pickers. Every control reports a shared `ShapeFillAction`.
	 */
	import type { ShapeFillAction, ShapeFillView } from 'ooxml-ui/pptx';

	import { useTranslator } from '../../../i18n/context';

	const {
		pattern,
		onaction,
		oncolorcommit,
	}: {
		pattern: NonNullable<ShapeFillView['pattern']>;
		onaction: (action: ShapeFillAction) => void;
		/** A settled colour pick, for the recent colours. */
		oncolorcommit: (hex: string) => void;
	} = $props();
	const t = useTranslator();

	function onColor(type: 'patternForeground' | 'patternBackground', hex: string): void {
		onaction({ type, hex });
		oncolorcommit(hex);
	}
</script>

<div class="pptx-svelte-pattern">
	<span class="pptx-svelte-pattern-label">{t('pptx.table.patternPreset')}</span>
	<div class="pptx-svelte-pattern-grid" role="listbox" aria-label={t('pptx.table.patternPreset')}>
		{#each pattern.swatches as swatch (swatch.value)}
			<button
				type="button"
				role="option"
				aria-selected={swatch.selected}
				data-testid="fill-pattern-swatch"
				class="pptx-svelte-pattern-swatch"
				class:pptx-svelte-pattern-swatch-on={swatch.selected}
				title={t(swatch.labelKey)}
				aria-label={t(swatch.labelKey)}
				onclick={() => onaction({ type: 'patternPreset', preset: swatch.value })}
			>
				{#if swatch.backgroundImage}
					<span
						class="pptx-svelte-pattern-swatch-fill"
						style={`background-image:${swatch.backgroundImage}`}
					></span>
				{/if}
			</button>
		{/each}
	</div>
	<label class="pptx-svelte-field">
		<span class="pptx-svelte-field-label">{t('pptx.fillAdvanced.foregroundColor')}</span>
		<input
			type="color"
			value={pattern.foreground}
			onchange={(e) => onColor('patternForeground', e.currentTarget.value)}
		/>
	</label>
	<label class="pptx-svelte-field">
		<span class="pptx-svelte-field-label">{t('pptx.fillAdvanced.backgroundColor')}</span>
		<input
			type="color"
			value={pattern.background}
			onchange={(e) => onColor('patternBackground', e.currentTarget.value)}
		/>
	</label>
</div>

<style>
	.pptx-svelte-pattern {
		display: flex;
		flex-direction: column;
		gap: 6px;
		margin-top: 10px;
	}

	.pptx-svelte-pattern-label {
		color: var(--pptx-muted-foreground, var(--office-muted-foreground, #6b7280));
	}

	.pptx-svelte-pattern-grid {
		display: grid;
		grid-template-columns: repeat(8, 1fr);
		gap: 4px;
		max-height: 160px;
		overflow-y: auto;
		padding: 4px;
		border: 1px solid var(--pptx-border, var(--office-border, #d1d5db));
		border-radius: 6px;
	}

	.pptx-svelte-pattern-swatch {
		display: flex;
		width: 100%;
		aspect-ratio: 1;
		padding: 0;
		border: 1px solid var(--pptx-border, var(--office-border, #d1d5db));
		border-radius: 4px;
		background: #fff;
		cursor: pointer;
	}

	.pptx-svelte-pattern-swatch-on {
		border-color: var(--pptx-primary, var(--office-accent, #2563eb));
		outline: 2px solid color-mix(in srgb, var(--pptx-primary, var(--office-accent, #2563eb)) 30%, transparent);
	}

	.pptx-svelte-pattern-swatch-fill {
		width: 100%;
		height: 100%;
		border-radius: 3px;
		background-repeat: repeat;
		background-size: 8px 8px;
	}

	.pptx-svelte-field {
		display: flex;
		flex-direction: row;
		align-items: center;
		justify-content: space-between;
		gap: 6px;
	}

	.pptx-svelte-field-label {
		color: var(--pptx-muted-foreground, var(--office-muted-foreground, #6b7280));
	}

	.pptx-svelte-field input[type='color'] {
		width: 40px;
		height: 26px;
		padding: 0;
		border: 1px solid var(--pptx-border, var(--office-border, #d1d5db));
		border-radius: 4px;
		background: transparent;
		cursor: pointer;
	}
</style>

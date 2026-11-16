<script lang="ts">
	/**
	 * GradientPanel: the gradient stop list of the Fill section, one row per
	 * shared `ShapeFillStopView` (colour with its theme grid, position,
	 * opacity, remove) and an Add Stop button. Every control reports a shared
	 * `ShapeFillAction`; `ShapeFillSection` turns it into the style patch. The
	 * gradient type and angle are shared descriptor fields drawn above it.
	 */
	import type { ShapeFillAction, ShapeFillStopView, ThemeColorPickerCommit } from 'ooxml-ui/pptx';

	import { useTranslator } from '../../../i18n/context';
	import ThemeColorSwatchGrid from './ThemeColorSwatchGrid.svelte';

	const {
		stops,
		themeColorMap,
		onaction,
		oncolorcommit,
	}: {
		stops: readonly ShapeFillStopView[];
		themeColorMap?: Record<string, string>;
		onaction: (action: ShapeFillAction) => void;
		/** A settled colour pick, for the recent colours. */
		oncolorcommit: (hex: string) => void;
	} = $props();
	const t = useTranslator();

	function onColor(index: number, hex: string): void {
		onaction({ type: 'stopColor', index, hex });
		oncolorcommit(hex);
	}
	function onTheme(index: number, commit: ThemeColorPickerCommit): void {
		onaction({ type: 'stopColor', index, hex: commit.hex, ref: commit.ref });
		oncolorcommit(commit.hex);
	}
</script>

<div class="pptx-svelte-gradient">
	<div class="pptx-svelte-gradient-stops">
		{#each stops as stop (stop.index)}
			<div class="pptx-svelte-gradient-stop-group" data-testid="fill-gradient-stop">
				<div class="pptx-svelte-gradient-stop">
					<input
						type="color"
						aria-label={t('pptx.gradient.stops')}
						value={stop.color}
						onchange={(e) => onColor(stop.index, e.currentTarget.value)}
					/>
					<input
						type="number"
						min="0"
						max="100"
						aria-label={t('pptx.gradient.position')}
						value={stop.position}
						onchange={(e) =>
							onaction({
								type: 'stopPosition',
								index: stop.index,
								position: Number(e.currentTarget.value),
							})}
					/>
					<input
						type="number"
						min="0"
						max="100"
						aria-label={t('pptx.strokeEffects.fillOpacity')}
						value={stop.opacity}
						onchange={(e) =>
							onaction({
								type: 'stopOpacity',
								index: stop.index,
								opacity: Number(e.currentTarget.value),
							})}
					/>
					<button
						type="button"
						class="pptx-svelte-gradient-remove"
						disabled={!stop.removable}
						aria-label={t('pptx.gradient.removeStop')}
						title={t('pptx.gradient.removeStop')}
						onclick={() => onaction({ type: 'removeStop', index: stop.index })}
					>
						&#10005;
					</button>
				</div>
				<ThemeColorSwatchGrid
					{themeColorMap}
					selectedRef={stop.colorRef}
					selectedHex={stop.color}
					onpick={(commit) => onTheme(stop.index, commit)}
				/>
			</div>
		{/each}
	</div>

	<button type="button" class="pptx-svelte-gradient-add" onclick={() => onaction({ type: 'addStop' })}>
		{t('pptx.gradient.addStop')}
	</button>
</div>

<style>
	.pptx-svelte-gradient {
		margin-top: 10px;
		padding-top: 10px;
		border-top: 1px solid var(--pptx-border, var(--office-border, #d1d5db));
		display: flex;
		flex-direction: column;
		gap: 8px;
	}

	.pptx-svelte-gradient-stops {
		display: flex;
		flex-direction: column;
		gap: 10px;
	}

	.pptx-svelte-gradient-stop-group {
		display: flex;
		flex-direction: column;
		gap: 4px;
	}

	.pptx-svelte-gradient-stop {
		display: grid;
		grid-template-columns: 32px 1fr 1fr 24px;
		align-items: center;
		gap: 6px;
	}

	.pptx-svelte-gradient-stop input[type='color'] {
		width: 32px;
		height: 24px;
		padding: 0;
		border: 1px solid var(--pptx-border, var(--office-border, #d1d5db));
		border-radius: 4px;
		background: transparent;
		cursor: pointer;
	}

	.pptx-svelte-gradient-stop input[type='number'] {
		height: 24px;
		box-sizing: border-box;
		padding: 0 6px;
		border: 1px solid var(--pptx-border, var(--office-border, #d1d5db));
		border-radius: var(--pptx-radius, var(--office-radius, 4px));
		background: var(--pptx-background, var(--office-background, #ffffff));
		color: inherit;
		font: inherit;
	}

	.pptx-svelte-gradient-remove {
		width: 24px;
		height: 24px;
		border: none;
		border-radius: 4px;
		background: transparent;
		color: inherit;
		cursor: pointer;
	}

	.pptx-svelte-gradient-remove:disabled {
		opacity: 0.35;
		cursor: default;
	}

	.pptx-svelte-gradient-add {
		height: 26px;
		border: 1px dashed var(--pptx-border, var(--office-border, #d1d5db));
		border-radius: var(--pptx-radius, var(--office-radius, 4px));
		background: transparent;
		color: inherit;
		cursor: pointer;
		font: inherit;
	}
</style>

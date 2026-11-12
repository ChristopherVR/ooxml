<script lang="ts">
	/**
	 * RibbonDisplayReveal: the slim strip that stays where the ribbon was in full-screen mode.
	 * Pressing it lays the ribbon back over the slide (the controller's `reveal`).
	 */
	import { RIBBON_DISPLAY_REVEAL_LABEL_KEY } from 'ooxml-ui/pptx';
	import type { RibbonDisplayController } from 'ooxml-ui/pptx';

	import { useTranslator } from '../../../i18n/context';
	import { controllerState } from '../../state/controller-state';

	const { ribbonDisplay }: { ribbonDisplay: RibbonDisplayController } = $props();
	const t = useTranslator();
	// svelte-ignore state_referenced_locally
	const state = controllerState(ribbonDisplay);
	const label = $derived(t(RIBBON_DISPLAY_REVEAL_LABEL_KEY));
</script>

<!-- The strip keeps its place while the ribbon is peeked over it, so the slide does not jump. -->
{#if state.current.effectiveMode === 'fullScreen'}
	<button
		type="button"
		class="pptx-svelte-ribbon-reveal"
		data-pptx-chrome="ribbon-reveal"
		aria-label={label}
		title={label}
		onclick={() => ribbonDisplay.reveal()}
	>
		···
	</button>
{/if}

<style>
	.pptx-svelte-ribbon-reveal {
		flex: none;
		display: flex;
		align-items: center;
		justify-content: center;
		width: 100%;
		height: 14px;
		padding: 0;
		border: none;
		border-bottom: 1px solid var(--pptx-border, var(--office-border, #d1d5db));
		background: var(--pptx-card, var(--office-background, #ffffff));
		color: var(--pptx-muted-foreground, var(--office-muted-foreground, #6b7280));
		font: inherit;
		font-size: 14px;
		line-height: 1;
		letter-spacing: 2px;
		cursor: pointer;
	}

	.pptx-svelte-ribbon-reveal:hover {
		background: var(--pptx-accent, var(--office-selected, #e5e7eb));
		color: var(--pptx-card-foreground, var(--office-foreground, #1f2937));
	}
</style>

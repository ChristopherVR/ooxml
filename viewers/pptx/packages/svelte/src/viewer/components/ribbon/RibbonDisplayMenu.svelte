<script lang="ts">
	/**
	 * RibbonDisplayMenu: the Ribbon Display Options menu, hung under the right end of the tab row.
	 * It draws the controller's `menu` as radio items; the controller owns the state and decides
	 * what the mode means. The anchor it sits in is zero-height, so the scrolling tab row does not
	 * clip it.
	 */
	import { RIBBON_DISPLAY_OPTIONS_LABEL_KEY } from 'ooxml-ui/pptx';
	import type { RibbonDisplayController } from 'ooxml-ui/pptx';

	import { useTranslator } from '../../../i18n/context';
	import { controllerState } from '../../state/controller-state';

	const { ribbonDisplay }: { ribbonDisplay: RibbonDisplayController } = $props();
	const t = useTranslator();
	// svelte-ignore state_referenced_locally
	const state = controllerState(ribbonDisplay);
</script>

{#if state.current.menuOpen}
	<div
		class="pptx-svelte-ribbon-display-menu"
		role="menu"
		aria-label={t(RIBBON_DISPLAY_OPTIONS_LABEL_KEY)}
		data-pptx-chrome="ribbon-display-menu"
	>
		{#each state.current.menu as item (item.mode)}
			<button
				type="button"
				class="pptx-svelte-ribbon-display-item"
				class:pptx-svelte-ribbon-display-item-checked={item.checked}
				role="menuitemradio"
				aria-checked={item.checked}
				data-ribbon-display={item.mode}
				onclick={() => ribbonDisplay.choose(item.mode)}
			>
				{t(item.labelKey)}
			</button>
		{/each}
	</div>
{/if}

<style>
	.pptx-svelte-ribbon-display-menu {
		position: absolute;
		top: 2px;
		right: 6px;
		z-index: 40;
		display: flex;
		flex-direction: column;
		min-width: 200px;
		padding: 4px 0;
		border: 1px solid var(--pptx-border, var(--office-border, #d1d5db));
		border-radius: 6px;
		background: var(--pptx-popover, var(--office-popover, var(--office-background, #ffffff)));
		color: var(--pptx-popover-foreground, var(--office-popover-foreground, var(--office-foreground, #1f2937)));
		box-shadow: 0 8px 20px rgb(0 0 0 / 0.2);
		font: inherit;
		font-size: 12px;
	}

	.pptx-svelte-ribbon-display-item {
		position: relative;
		padding: 6px 14px 6px 26px;
		border: none;
		background: transparent;
		color: inherit;
		font: inherit;
		text-align: left;
		white-space: nowrap;
		cursor: pointer;
	}

	.pptx-svelte-ribbon-display-item:hover {
		background: var(--pptx-accent, var(--office-selected, #e5e7eb));
	}

	.pptx-svelte-ribbon-display-item-checked::before {
		content: '';
		position: absolute;
		left: 9px;
		top: 50%;
		width: 9px;
		height: 5px;
		border-left: 1.5px solid currentColor;
		border-bottom: 1.5px solid currentColor;
		transform: translateY(-70%) rotate(-45deg);
	}
</style>

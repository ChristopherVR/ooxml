<script lang="ts">
	/**
	 * BuildingBlocksToolbar: the compact toolbar block. One button per control of the shared
	 * descriptor (`describeBuildingBlockToolbar`), grouped in `data-group` spans; insert controls
	 * raise a building-block request instead of opening a picker or dialog.
	 */
	import type { BuildingBlockToolbarGroupId } from 'ooxml-ui/pptx';

	import type { ViewerBuildingBlocks } from './types';

	const { blocks, include, class: className = '' }: {
		blocks: ViewerBuildingBlocks;
		/** Show only these groups, in descriptor order. */
		include?: readonly BuildingBlockToolbarGroupId[] | undefined;
		class?: string;
	} = $props();

	const groups = $derived(
		blocks.toolbar.filter((group) => !include || include.includes(group.id)),
	);
</script>

<div class={`pptx-svelte-bb-toolbar ${className}`} data-pptx-block="toolbar" role="toolbar">
	{#each groups as group (group.id)}
		<span data-group={group.id} role="group">
			{#each group.controls as control (control.id)}
				<button
					type="button"
					data-control={control.id}
					disabled={control.disabled}
					aria-pressed={control.pressed === undefined ? undefined : control.pressed}
					onclick={() => blocks.runToolbarControl(control.id)}
				>{blocks.t(control.labelKey)}</button>
			{/each}
		</span>
	{/each}
</div>

<style>
	.pptx-svelte-bb-toolbar {
		display: flex;
		flex-wrap: wrap;
		gap: 8px;
		align-items: center;
	}

	.pptx-svelte-bb-toolbar span {
		display: inline-flex;
		gap: 4px;
	}
</style>

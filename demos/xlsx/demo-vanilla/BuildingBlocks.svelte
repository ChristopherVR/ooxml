<script lang="ts">
 import { createViewerBuildingBlocks, type BuildingBlockOptions } from '../../../viewers/xlsx/packages/svelte/src/runtime';
 import SpreadsheetCanvas from '../../../viewers/xlsx/packages/svelte/src/SpreadsheetCanvas.svelte';
 import Toolbar from '../../../viewers/xlsx/packages/svelte/src/Toolbar.svelte';
 import FormulaBar from '../../../viewers/xlsx/packages/svelte/src/FormulaBar.svelte';
 import SheetTabs from '../../../viewers/xlsx/packages/svelte/src/SheetTabs.svelte';
 let { initial }: { initial: BuildingBlockOptions } = $props();
 let options = $state<BuildingBlockOptions>({});
 $effect.pre(() => { options = initial; });
 const blocks = createViewerBuildingBlocks();
 $effect(() => blocks.update(options));
 export function update(next: BuildingBlockOptions) { options = next; }
 export const handle = blocks.handle;
 export function snapshot() { return $blocks; }
</script>
{#if $blocks.toolbarProps}
 <Toolbar toolbarProps={$blocks.toolbarProps} />
 <button type="button" onclick={() => $blocks.toolbarProps?.onItalic()}>Host italic</button>
 <FormulaBar toolbarProps={$blocks.toolbarProps} /><SheetTabs toolbarProps={$blocks.toolbarProps} />
{/if}
<SpreadsheetCanvas {...blocks.canvasProps} />

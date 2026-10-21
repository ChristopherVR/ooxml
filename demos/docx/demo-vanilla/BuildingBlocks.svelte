<script lang="ts">
 import { writable } from 'svelte/store';
 import { createViewerBuildingBlocks, type BuildingBlockOptions } from '../../../viewers/docx/packages/svelte/src/runtime';
 import DocumentCanvas from '../../../viewers/docx/packages/svelte/src/DocumentCanvas.svelte';
 import Toolbar from '../../../viewers/docx/packages/svelte/src/Toolbar.svelte';
 import type { NativeHarness } from './building-blocks-types';
 let { initial, ready }: { initial: BuildingBlockOptions; ready: (harness: Omit<NativeHarness, 'destroy'>) => void } = $props();
 const options = writable((() => initial)());
 const blocks = createViewerBuildingBlocks(options);
 const { toolbarProps, error, mode } = blocks;
 (() => ready)()({ handle: blocks.handle, toolbar: () => $toolbarProps, error: () => $error, mode: () => $mode,
  permission: (canEdit) => options.update((current) => ({ ...current, canEdit })), options: (next) => options.set(next) });
</script>
<DocumentCanvas {...blocks.canvasProps} className={$mode} />
{#if $toolbarProps}<Toolbar toolbarProps={$toolbarProps} />{/if}

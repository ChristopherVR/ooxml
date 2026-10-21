<script lang="ts">
 import { untrack } from 'svelte';
 import type { Readable } from 'svelte/store';
 import { createViewerBuildingBlocks, type BuildingBlockOptions } from './building-blocks';
 import DiagramCanvas from './DiagramCanvas.svelte';
 import Toolbar from './Toolbar.svelte';
 import PageNavigation from './PageNavigation.svelte';
 import Inspector from './Inspector.svelte';
 import ShapesPane from './ShapesPane.svelte';
 let { options, ready }: { options: Readable<BuildingBlockOptions>; ready(blocks: ReturnType<typeof createViewerBuildingBlocks>): void } = $props();
 const blocks = createViewerBuildingBlocks(untrack(() => options)); const toolbar = blocks.toolbarProps; const mode = blocks.mode; const shapes = blocks.shapesProps;
 untrack(() => ready(blocks));
</script>
<DiagramCanvas {...blocks.canvasProps} /><Toolbar toolbarProps={$toolbar} /><PageNavigation toolbarProps={$toolbar} /><Inspector {...blocks.inspectorProps} /><ShapesPane {...$shapes} /><output>{$mode}:{$toolbar.zoom}</output>

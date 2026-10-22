<script lang="ts">
 import Controls from './Controls.svelte';
 import { toolbarControls, type WordToolbarProps } from './runtime';
 let { toolbarProps }: { toolbarProps: WordToolbarProps } = $props();
</script>
<Controls role="toolbar" groups={toolbarControls(toolbarProps)} />

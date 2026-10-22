<script lang="ts">
 import Controls from './Controls.svelte';
 import { statusControls, type WordStatusProps } from './runtime';
 let { statusProps }: { statusProps: WordStatusProps } = $props();
</script>
<Controls role="group" groups={statusControls(statusProps)} />

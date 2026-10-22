<script lang="ts">
 import Controls from './Controls.svelte';
 import { navigationControls, type WordNavigationProps } from './runtime';
 let { navigationProps }: { navigationProps: WordNavigationProps } = $props();
</script>
<Controls role="group" groups={navigationControls(navigationProps)} />

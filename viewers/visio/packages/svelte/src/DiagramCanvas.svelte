<script lang="ts">
 import { onMount } from 'svelte';
 let { mount, class: className }: { mount(node: HTMLElement | null): void; class?: string } = $props();
 let host: HTMLDivElement;
 onMount(() => { mount(host); return () => mount(null); });
</script>
<div bind:this={host} class={className}></div>

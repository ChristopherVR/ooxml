<script lang="ts">
 let { mount, class: className }: { mount(node: HTMLElement | null): void; class?: string } = $props();
 let host: HTMLDivElement;
 $effect(() => { const current = mount; current(host); return () => current(null); });
</script>
<div bind:this={host} class={className}></div>

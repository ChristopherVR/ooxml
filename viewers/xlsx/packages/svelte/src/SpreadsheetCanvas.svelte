<script lang="ts">
 let { mount, class: className }: { mount: (host: HTMLElement | null) => void; class?: string } = $props();
 let host = $state<HTMLDivElement>();
 $effect(() => { if (!host) return; const current = mount; current(host); return () => current(null); });
</script>
<div bind:this={host} class={className}></div>

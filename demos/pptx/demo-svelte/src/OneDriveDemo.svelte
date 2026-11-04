<script lang="ts">
	import { PowerPointViewer } from 'pptx-svelte-viewer';
	import { onMount } from 'svelte';

	import { mountOneDriveDemo } from '../../shared/onedrive-demo';
	import type { OneDriveDemoDeck, OneDriveDemoMount } from '../../shared/onedrive-demo';

	// `?onedrive=1`: open, edit and save a deck on your own OneDrive (docs/onedrive-quickstart.md).
	let bar: HTMLDivElement;
	let deck = $state.raw<OneDriveDemoDeck | null>(null);
	let storage = $state.raw<OneDriveDemoMount['storage']>();
	onMount(() => {
		const mounted = mountOneDriveDemo(bar, (next) => { deck = next; });
		storage = mounted.storage;
		return () => mounted.dispose();
	});
</script>

<div bind:this={bar}></div>
<main class="od-viewer">
	{#if deck && storage}
		<PowerPointViewer source={deck.bytes} fileName={deck.fileName} editable={deck.canEdit} {storage} />
	{:else}
		<p class="od-empty">No OneDrive file open yet.</p>
	{/if}
</main>

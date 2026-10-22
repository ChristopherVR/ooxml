<script lang="ts">
	import { PowerPointViewer } from 'pptx-svelte-viewer';
	import { onMount } from 'svelte';

	import { createSameBrowserDemo } from '../../shared/same-browser-collaboration';
	import type { SameBrowserDemo } from '../../shared/same-browser-collaboration';

	let demo = $state.raw<SameBrowserDemo | null>(null);
	let error = $state('');
	onMount(() => {
		let disposed = false;
		let current: SameBrowserDemo | undefined;
		void createSameBrowserDemo().then((value) => {
			if (disposed) { value.dispose(); return; }
			current = value;
			demo = value;
			return undefined;
		}).catch((reason: unknown) => { error = String(reason); });
		return () => { disposed = true; current?.dispose(); };
	});
</script>

<main style:position="fixed" style:inset="28px 0 0">
	{#if error}<p role="alert">{error}</p>{/if}
	{#if demo}
		<PowerPointViewer source={demo.source} fileName={demo.fileName} collaboration={demo.collaboration} editable={true} oncollaborationchange={demo.report} />
	{/if}
</main>
